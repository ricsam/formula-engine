import { Menu, X } from 'lucide-react'

interface MobileMenuToggleProps {
  isOpen: boolean
  onToggle: () => void
}

export function MobileMenuToggle({ isOpen, onToggle }: MobileMenuToggleProps) {
  return (
    <button
      onClick={onToggle}
      className="lg:hidden fixed top-4 left-4 z-50 p-2 bg-white border rounded-lg shadow-md hover:bg-gray-50 transition-colors"
      aria-label={isOpen ? 'Close menu' : 'Open menu'}
    >
      {isOpen ? (
        <X className="w-6 h-6 text-gray-600" />
      ) : (
        <Menu className="w-6 h-6 text-gray-600" />
      )}
    </button>
  )
}
