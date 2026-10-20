import { useState } from 'react'
import { OctagonX, Play, Square } from 'lucide-react'
import { Button, Card, Checkbox, Modal, ProgressBar } from '../ui'
import { useStore } from '../../state/store'
import { STAGE_LABELS } from '../../types'
import { formatElapsed } from '../../lib/utils'

export function ProcessingCard() {
  const s = useStore()
  const { run } = s
  const running = s.running
  const stopping = run.status === 'stopping'
  const cancelling = run.status === 'cancelling'
  const [confirmCancel, setConfirmCancel] = useState(false)

  const headline =
    run.status === 'processing'
      ? `Đang xử lý: ${run.currentFile}`
      : cancelling
        ? `Đang dừng — báo cáo hiện tại dừng ở điểm an toàn gần nhất: ${run.currentFile}`
        : stopping
          ? `Đang chờ dừng — kết thúc báo cáo hiện tại: ${run.currentFile}`
          : run.status === 'done'
            ? run.cancelRequested
              ? `Đã dừng toàn bộ theo yêu cầu — ${run.doneCount}/${run.queue.length} báo cáo đã hoàn tất`
              : run.stopped
                ? `Đã dừng theo yêu cầu — ${run.doneCount}/${run.queue.length} báo cáo đã hoàn tất`
                : `Đã kết thúc lượt xử lý — ${run.doneCount}/${run.queue.length} báo cáo`
            : s.queue.length
              ? `Sẵn sàng xử lý ${s.queue.length} báo cáo trong hàng đợi`
              : 'Sẵn sàng — chưa có báo cáo nào trong hàng đợi'

  const stageText = running
    ? STAGE_LABELS[run.stage]
    : run.status === 'done'
      ? run.cancelRequested
        ? 'Đã dừng toàn bộ'
        : run.stopped
          ? 'Đã dừng an toàn'
          : 'Kết thúc'
      : 'Chưa bắt đầu'

  return (
    <Card title="Xử lý">
      <div className="flex flex-wrap items-center gap-2">
        <Button
          variant="primary"
          className="h-[30px] px-3"
          icon={<Play className="h-3.5 w-3.5" />}
          disabled={running || s.locked || s.scanDirty || !s.scanned || s.queue.length === 0}
          onClick={s.startProcessing}
        >
          Bắt đầu xử lý
        </Button>
        <Button
          variant="danger"
          className="h-[30px] px-3"
          icon={<Square className="h-3.5 w-3.5" />}
          disabled={!running || stopping || cancelling}
          onClick={s.stopAfterCurrent}
        >
          {stopping ? 'Đang chờ dừng sau file hiện tại…' : 'Dừng sau file hiện tại'}
        </Button>
        {/* PROMPT-024R: cancel-all requires confirmation; duplicate clicks are disabled once requested. */}
        <Button
          variant="danger"
          className="h-[30px] px-3"
          icon={<OctagonX className="h-3.5 w-3.5" />}
          disabled={!running || cancelling}
          onClick={() => setConfirmCancel(true)}
        >
          {cancelling ? 'Đang dừng…' : 'Dừng tất cả'}
        </Button>
        <Checkbox
          className="ml-3"
          checked={s.force}
          disabled={running}
          onChange={s.setForce}
          label="Xử lý lại báo cáo đã xử lý — ghi đè các trường tự động"
        />
        <span className="ml-auto text-xs2 text-muted">
          {running ? 'Đang xử lý — các thao tác thay đổi nguồn và hàng đợi đã bị vô hiệu hóa.' : `Chế độ: ${s.aiMode === 'model' ? 'mô hình ' + s.ollama.model : 'phân tích từ khoá (heuristic)'}`}
        </span>
      </div>

      <div className="mt-2 border-t border-lineSoft pt-2">
        <div className="flex items-center gap-2 text-base2">
          <span className="min-w-0 flex-1 truncate font-medium text-header">{headline}</span>
          <span className="text-[19px] font-semibold leading-[22px] text-brand-600">{Math.round(run.percent)}%</span>
        </div>
        <ProgressBar className="mt-1.5" value={run.percent} />
        <div className="mt-1.5 flex flex-wrap items-center gap-x-5 gap-y-1 text-xs2">
          <span className="text-body">
            Giai đoạn: <span className="font-medium">{stageText}</span>
          </span>
          <span className="text-body">
            File hiện tại: <span className="font-mono">{running ? run.currentFile : '—'}</span>
          </span>
          <span className="text-body">
            Đã chạy: <span className="font-mono">{formatElapsed(run.elapsedSec)}</span>
          </span>
          <span className="text-body">
            Còn khoảng:{' '}
            <span className="font-mono">
              {run.remainSec === null ? '— (chưa đủ dữ liệu)' : formatElapsed(run.remainSec)}
            </span>
          </span>
          {typeof run.cancelledCount === 'number' && run.cancelledCount > 0 && (
            <span className="text-body">
              Đã hủy: <span className="font-mono">{run.cancelledCount}</span>
            </span>
          )}
          <span className="text-muted">
            Tiến độ tính từ giai đoạn thật của từng báo cáo, không nội suy theo thời gian.
          </span>
        </div>
      </div>

      <Modal
        open={confirmCancel}
        title="Dừng toàn bộ xử lý?"
        width="max-w-md"
        onClose={() => setConfirmCancel(false)}
        footer={
          <>
            <Button className="mr-auto" onClick={() => setConfirmCancel(false)}>
              Tiếp tục xử lý
            </Button>
            <Button
              variant="danger"
              icon={<OctagonX className="h-3.5 w-3.5" />}
              disabled={cancelling}
              onClick={() => {
                setConfirmCancel(false)
                void s.cancelAll()
              }}
            >
              Dừng tất cả
            </Button>
          </>
        }
      >
        <p className="text-base2 text-body">
          Báo cáo đang chạy sẽ được dừng tại điểm an toàn gần nhất.
          Các báo cáo còn lại sẽ không được xử lý.
        </p>
        <p className="mt-2 text-xs2 text-muted">
          Kết quả đã ghi Excel trước đó được giữ nguyên; báo cáo đang chạy dở sẽ hiển thị “Đã hủy”
          chứ không bị tính là lỗi.
        </p>
      </Modal>
    </Card>
  )
}
