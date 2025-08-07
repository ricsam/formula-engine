import { createFileRoute } from '@tanstack/react-router'
import { MultiSheetDemo } from '../MultiSheetDemo'

export const Route = createFileRoute('/multisheet')({
  component: MultisheetPage,
})

function MultisheetPage() {
  return <MultiSheetDemo />
}
