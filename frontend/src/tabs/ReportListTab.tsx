import { FooterBar } from '../components/report/FooterBar'
import { LogCard, StatsCard } from '../components/report/LogAndStats'
import { ProcessingCard } from '../components/report/ProcessingCard'
import { ReportDetail } from '../components/report/ReportDetail'
import { ReportTable } from '../components/report/ReportTable'
import { SourcePanel } from '../components/report/SourcePanel'

export function ReportListTab() {
  return (
    <div className="flex min-h-full flex-col gap-3 p-3">
      <SourcePanel />

      {/* Khu vực làm việc: bảng báo cáo 65% — khung chi tiết 35% */}
      <div className="grid grid-cols-1 items-stretch gap-3 xl:grid-cols-[65fr_35fr]">
        <div className="flex min-h-[430px] flex-col">
          <ReportTable />
        </div>
        <div className="min-h-0">
          <ReportDetail />
        </div>
      </div>

      <ProcessingCard />

      <div className="grid grid-cols-1 gap-3 xl:grid-cols-[65fr_35fr]">
        <LogCard />
        <StatsCard />
      </div>

      <FooterBar />
    </div>
  )
}
