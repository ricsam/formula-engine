import { createFileRoute } from '@tanstack/react-router'
import { EventsDemo } from '../EventsDemo'

export const Route = createFileRoute('/events')({
  component: EventsPage,
})

function EventsPage() {
  return <EventsDemo />
}
