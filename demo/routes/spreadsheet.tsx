import { createFileRoute } from '@tanstack/react-router'
import { FullSpreadsheetDemo } from '../FullSpreadsheetDemo'

export const Route = createFileRoute('/spreadsheet')({
  component: SpreadsheetPage,
})

function SpreadsheetPage() {
  return <FullSpreadsheetDemo />
}
