import { createFileRoute } from '@tanstack/react-router'
import { ExcelDemo } from '../ExcelDemo'

export const Route = createFileRoute('/excel')({
  component: ExcelPage,
})

function ExcelPage() {
  return <ExcelDemo />
}
