"""Durable review evidence, separate from reusable classification/label identities.

Only learning_data/previews is authoritative. Source paths live in private snapshots,
never DTOs. Rendered images are content-addressed; atomic metadata replacement commits
an image+geometry pair. A corrupt/missing image must never remove its label record.
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import logging
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Optional

from PIL import Image

from .image_learning import ImageCandidate
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.learning_assets")
SCHEMA = 1
_IMAGE_NAME = re.compile(r"[a-f0-9]{32}\.(?:jpg|png)\Z")
_TOKEN = re.compile(r"[a-f0-9]{24}\Z")


def digest(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()[:24]


def asset_key(candidate) -> str:
    return digest((candidate.candidate_id, report_scope_key(candidate.source_file)))


def review_id(candidate) -> str:
    return "review:" + asset_key(candidate)


def source_version(source: str):
    try:
        stat = Path(source).stat() if source else None
        return [stat.st_mtime_ns, stat.st_size] if stat is not None else None
    except OSError:
        return None


def source_name(name: str) -> str:
    return str(name or "").replace("\\", "/").rsplit("/", 1)[-1]


class LearningAssets:
    def __init__(self, directory: Path):
        self.root = Path(directory) / "previews"

    @staticmethod
    def _read(path: Path) -> dict:
        try:
            if path.stat().st_size > 4 * 1024 * 1024:
                return {}
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) and value.get("schema") == SCHEMA else {}
        except (OSError, ValueError, TypeError):
            return {}

    @staticmethod
    def _atomic(path: Path, data: bytes) -> None:
        try:
            if path.is_file() and path.read_bytes() == data:
                return                         # no duplicate writes on every Learning poll
        except OSError:
            pass
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = None
        try:
            with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".tmp", delete=False) as stream:
                temp = Path(stream.name)
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
            temp.replace(path)
        finally:
            if temp is not None:
                temp.unlink(missing_ok=True)

    def _json(self, path: Path, value: dict) -> None:
        self._atomic(path, json.dumps(value, ensure_ascii=False, sort_keys=True).encode("utf-8"))

    def _put_image(self, src: str) -> str:
        if not src or len(src) > 44 * 1024 * 1024:
            return ""
        raw = base64.b64decode(src.split(",", 1)[1], validate=True)
        with Image.open(io.BytesIO(raw)) as image:
            ext = {"JPEG": "jpg", "PNG": "png"}.get(image.format)
            image.verify()
        if ext is None:
            return ""
        name = hashlib.sha256(raw).hexdigest()[:32] + "." + ext
        self._atomic(self.root / name, raw)
        return name

    def _get_image(self, name: str) -> tuple:
        if not isinstance(name, str) or not _IMAGE_NAME.fullmatch(name):
            return "", 0, 0
        try:
            path = self.root / name
            if path.stat().st_size > 32 * 1024 * 1024:
                return "", 0, 0
            raw = path.read_bytes()
            if hashlib.sha256(raw).hexdigest()[:32] != name.split(".")[0]:
                return "", 0, 0
            with Image.open(io.BytesIO(raw)) as image:
                width, height = image.size
                image.verify()
            mime = "jpeg" if name.endswith(".jpg") else "png"
            return "data:image/" + mime + ";base64," + base64.b64encode(raw).decode("ascii"), width, height
        except (OSError, ValueError, Image.DecompressionBombError):
            return "", 0, 0

    def snapshot(self, candidate) -> dict:
        value = self._read(self.root / (asset_key(candidate) + "_candidate.json"))
        fields = value.get("candidate")
        return value if isinstance(fields, dict) and fields.get("candidate_id") == candidate.candidate_id else {}

    def remember_candidate(self, candidate, src: str, preview: Optional[dict] = None) -> None:
        try:
            previous = self.snapshot(candidate)
            version = source_version(candidate.source_file)
            if version is None and previous:
                version = previous.get("version")
            fields = {key: getattr(candidate, key) for key in ImageCandidate.__dataclass_fields__
                      if hasattr(candidate, key)}
            fields["source_name"] = source_name(candidate.source_name)
            fields["thumbnail"] = ""           # no session/temp thumbnail path is authoritative
            image = self._put_image(src) if src else ""
            if not image and previous.get("version") == version:
                image = previous.get("image", "")
            value = {"schema": SCHEMA, "candidate": fields, "version": version, "image": image,
                     "slide": (preview or {}).get("_asset_key", previous.get("slide", ""))}
            self._json(self.root / (asset_key(candidate) + "_candidate.json"), value)
        except Exception as exc:  # noqa: BLE001 – optional display persistence cannot stop extraction
            LOG.warning("LEARNING_ASSET_WRITE_FAILED stage=candidate exception_type=%s", type(exc).__name__)

    def candidate_src(self, candidate) -> str:
        if getattr(candidate, "_snapshot_changed", False):
            return ""
        value = self.snapshot(candidate)
        current = source_version(candidate.source_file)
        if current is not None and current != value.get("version"):
            return ""                          # a live source changed; don't pin an old candidate bitmap
        src = self._get_image(value.get("image", ""))[0]
        if src or candidate.source_file:
            return src
        # Legacy source-less labels may use their old thumbnail, never an arbitrary glob match.
        name = hashlib.sha1(candidate.candidate_id.encode("utf-8")).hexdigest()[:16] + ".png"
        try:
            raw = (self.root.parent / "thumbnails" / name).read_bytes()
            with Image.open(io.BytesIO(raw)) as image:
                image.verify()
            return "data:image/png;base64," + base64.b64encode(raw).decode("ascii")
        except (OSError, ValueError):
            return ""

    def candidates(self, records: dict) -> list:
        out = []
        for path in sorted(self.root.glob("*_candidate.json")):
            value = self._read(path)
            try:
                fields = {key: val for key, val in value["candidate"].items()
                          if key in ImageCandidate.__dataclass_fields__}
                candidate = ImageCandidate(**fields)
                if (not isinstance(candidate.source_file, str) or not isinstance(candidate.candidate_id, str)
                    or not isinstance(candidate.features, dict) or not isinstance(candidate.evidence, list)
                    or any(not isinstance(v, str) for v in candidate.evidence)
                    or any(len(getattr(candidate, k)) != n for k, n in
                           (("bounds", 4), ("block_bounds", 4), ("slide_size", 2)))):
                    continue
                for k in ("bounds", "block_bounds", "slide_size"):
                    setattr(candidate, k, tuple(int(v) for v in getattr(candidate, k)))
                candidate.slide, candidate.picture_id, candidate.order = int(candidate.slide), int(candidate.picture_id), int(candidate.order)
                candidate.confidence = float(candidate.confidence)
                candidate.item_index = int(candidate.item_index)
                if path.name != asset_key(candidate) + "_candidate.json" or candidate.hard_excluded:
                    continue
                candidate.source_name = source_name(candidate.source_name)
                current = source_version(candidate.source_file)
                candidate._snapshot_version = value.get("version")
                candidate._snapshot_changed = current is not None and current != candidate._snapshot_version
                if candidate._snapshot_changed:
                    candidate.excel_output_eligible = False
                    candidate.eligibility_reason = "Nguồn đã thay đổi; chạy xử lý lại để cập nhật preview."
                candidate._persisted = True
                out.append(candidate)
            except (KeyError, TypeError, ValueError, AttributeError):
                continue
        known = {c.candidate_id for c in out}
        for cid, record in records.items():
            if cid in known:
                continue
            # An old label without provenance still appears. Never guess a source path, bbox or eligibility.
            try:
                candidate = ImageCandidate(
                    cid, str(record.get("management_number") or ""), source_name(record.get("source_name", "")), "",
                    int(record.get("slide") or 0), int(record.get("shape_id") or record.get("picture_id") or 0), 0,
                    (0, 0, 0, 0), (1, 1), (0, 0, 0, 0), dict(record.get("features") or {}),
                    confidence=float(record.get("deterministic_confidence") or 0.5),
                    deterministic_kind=str(record.get("deterministic_kind") or "ambiguous"),
                    eligibility_reason="Bản ghi cũ không có hình học nguồn; chạy xử lý lại để xem preview.")
                candidate._persisted = True
                out.append(candidate)
            except (TypeError, ValueError):
                continue
        return sorted(out, key=lambda c: (report_scope_key(c.source_file), c.slide, c.order, c.picture_id, c.candidate_id))

    def remember_slide(self, key: tuple, result: dict) -> dict:
        try:
            token = digest(key)
            image = self._put_image(result["src"])
            if not image:
                return result
            metadata = {k: v for k, v in result.items() if k not in {"src", "created", "_asset_key"}}
            self._json(self.root / (token + "_slide.json"), {
                "schema": SCHEMA, "signature": digest(key[5:]), "image": image,
                "saved_at": time.time(), "metadata": metadata})
            return {**result, "_asset_key": token}
        except Exception as exc:  # noqa: BLE001
            LOG.warning("LEARNING_ASSET_WRITE_FAILED stage=slide exception_type=%s", type(exc).__name__)
            return result

    def slide(self, token: str, signature: tuple, *, offline: bool = False) -> dict:
        if not isinstance(token, str) or not _TOKEN.fullmatch(token):
            return {}
        value = self._read(self.root / (token + "_slide.json"))
        same = value.get("signature") == digest(signature)
        if not same and not offline:
            return {}
        src, width, height = self._get_image(value.get("image", ""))
        if not src:
            return {}
        try:
            metadata = value["metadata"]
            if not isinstance(metadata, dict) or metadata.get("backend") not in {"powerpoint", "libreoffice", "builtin"}:
                return {}
            for region in metadata.get("regions", []):
                if not isinstance(region, dict) or not all(isinstance(region.get("bbox", {}).get(k), int)
                                                           for k in ("x", "y", "width", "height")):
                    return {}
                if not isinstance(region.get("pictureShapeIds"), list):
                    return {}
            result = {**metadata, "src": src, "width": width, "height": height, "_asset_key": token,
                      "created": time.monotonic() - max(0, time.time() - float(value["saved_at"]))}
            if offline and not same:
                result["regions"] = []       # cannot recompute a new production decision without PPTX
            return result
        except (KeyError, TypeError, ValueError):
            return {}

    def offline_slide(self, candidate, signature: tuple) -> dict:
        if getattr(candidate, "_snapshot_changed", False):
            return {}
        return self.slide(self.snapshot(candidate).get("slide", ""), signature, offline=True)
