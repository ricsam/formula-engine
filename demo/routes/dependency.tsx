import { createFileRoute } from '@tanstack/react-router'
import { DependencyFlowDemo } from '../DependencyFlowDemo'

export const Route = createFileRoute('/dependency')({
  component: DependencyFlowDemo,
})
