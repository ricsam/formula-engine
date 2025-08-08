import { Link, useRouterState } from '@tanstack/react-router'
import { 
  Home, 
  Activity, 
  Grid3x3, 
  Layers, 
  Info,
  ChevronDown,
  ChevronRight,
  Calculator,
  BarChart3,
  Filter,
  Type,
  Search,
  Database,
  FileSpreadsheet,
  Network
} from 'lucide-react'
import { useState } from 'react'

interface NavigationItem {
  id: string
  label: string
  to?: string
  icon?: React.ReactNode
  children?: NavigationItem[]
}

const navigationData: NavigationItem[] = [
  {
    id: 'home',
    label: 'Home',
    to: '/',
    icon: <Home className="w-4 h-4" />
  },
  {
    id: 'demos',
    label: 'Demos',
    icon: <Grid3x3 className="w-4 h-4" />,
    children: [
      {
        id: 'spreadsheet',
        label: 'Full Spreadsheet',
        to: '/spreadsheet',
        icon: <Calculator className="w-4 h-4" />
      },
      {
        id: 'multisheet',
        label: 'Multi-Sheet',
        to: '/multisheet',
        icon: <Layers className="w-4 h-4" />
      },
      {
        id: 'excel',
        label: 'Excel Clone',
        to: '/excel',
        icon: <FileSpreadsheet className="w-4 h-4" />
      },
      {
        id: 'dependency',
        label: 'Dependency Graph',
        to: '/dependency',
        icon: <Network className="w-4 h-4" />
      }
    ]
  },
  {
    id: 'functions',
    label: 'Function Categories',
    icon: <BarChart3 className="w-4 h-4" />,
    children: [
      {
        id: 'math',
        label: 'Mathematical Functions',
        icon: <Calculator className="w-4 h-4" />,
        children: [
          { id: 'basic-math', label: 'Basic Math (ADD, SUBTRACT, etc.)' },
          { id: 'advanced-math', label: 'Advanced Math (SIN, COS, LOG, etc.)' },
          { id: 'statistical', label: 'Statistical (SUM, COUNT, AVERAGE, etc.)' }
        ]
      },
      {
        id: 'logical',
        label: 'Logical Functions',
        icon: <Grid3x3 className="w-4 h-4" />,
        children: [
          { id: 'conditions', label: 'Conditions (IF, AND, OR, etc.)' },
          { id: 'comparisons', label: 'Comparisons (EQ, LT, GT, etc.)' }
        ]
      },
      {
        id: 'text',
        label: 'Text Functions',
        icon: <Type className="w-4 h-4" />,
        children: [
          { id: 'text-funcs', label: 'String Functions (CONCATENATE, LEN, etc.)' }
        ]
      },
      {
        id: 'lookup',
        label: 'Lookup Functions',
        icon: <Search className="w-4 h-4" />,
        children: [
          { id: 'lookup-funcs', label: 'Lookup (VLOOKUP, INDEX, MATCH, etc.)' }
        ]
      },
      {
        id: 'array',
        label: 'Array Functions',
        icon: <Filter className="w-4 h-4" />,
        children: [
          { id: 'array-funcs', label: 'Array (FILTER, ARRAY_CONSTRAIN, etc.)' }
        ]
      },
      {
        id: 'info',
        label: 'Info Functions',
        icon: <Database className="w-4 h-4" />,
        children: [
          { id: 'info-funcs', label: 'Info (ISBLANK, ISERROR, etc.)' }
        ]
      }
    ]
  },
  {
    id: 'about',
    label: 'About',
    to: '/about',
    icon: <Info className="w-4 h-4" />
  }
]

interface NavigationItemProps {
  item: NavigationItem
  level: number
  currentPath: string
  onItemClick?: () => void
}

function NavigationItemComponent({ item, level, currentPath, onItemClick }: NavigationItemProps) {
  const [isExpanded, setIsExpanded] = useState(() => {
    // Auto-expand if current path matches or is a child
    if (item.to && currentPath === item.to) return true
    if (item.children) {
      return item.children.some(child => 
        child.to && currentPath === child.to ||
        child.children?.some(grandchild => grandchild.to && currentPath === grandchild.to)
      )
    }
    return false
  })

  const hasChildren = item.children && item.children.length > 0
  const isActive = item.to && currentPath === item.to
  const indentClass = level === 0 ? '' : level === 1 ? 'ml-4' : 'ml-8'

  const toggleExpanded = () => {
    if (hasChildren) {
      setIsExpanded(!isExpanded)
    }
  }

  const ItemContent = (
    <div 
      className={`
        flex items-center gap-2 px-3 py-2 rounded-lg cursor-pointer transition-colors
        ${isActive 
          ? 'bg-blue-100 text-blue-700 font-medium' 
          : 'hover:bg-gray-100 text-gray-700'
        }
        ${indentClass}
      `}
    >
      {hasChildren && (
        <div className="w-4 h-4 flex items-center justify-center">
          {isExpanded ? (
            <ChevronDown className="w-3 h-3" />
          ) : (
            <ChevronRight className="w-3 h-3" />
          )}
        </div>
      )}
      {!hasChildren && <div className="w-4" />}
      
      {item.icon && (
        <div className="flex-shrink-0">
          {item.icon}
        </div>
      )}
      
      <span className="text-sm truncate">{item.label}</span>
    </div>
  )

  const handleItemClick = () => {
    if (hasChildren) {
      toggleExpanded()
    } else if (item.to && onItemClick) {
      onItemClick()
    }
  }

  return (
    <div>
      {item.to ? (
        <Link to={item.to} className="block" onClick={onItemClick}>
          {ItemContent}
        </Link>
      ) : (
        <div onClick={handleItemClick}>
          {ItemContent}
        </div>
      )}
      
      {hasChildren && isExpanded && (
        <div className="mt-1 space-y-1">
          {item.children!.map((child) => (
            <NavigationItemComponent
              key={child.id}
              item={child}
              level={level + 1}
              currentPath={currentPath}
              onItemClick={onItemClick}
            />
          ))}
        </div>
      )}
    </div>
  )
}

interface SidebarProps {
  isOpen?: boolean
  onClose?: () => void
}

export function Sidebar({ isOpen = true, onClose }: SidebarProps) {
  const router = useRouterState()
  const currentPath = router.location.pathname

  return (
    <>
      {/* Mobile Overlay */}
      {isOpen && (
        <div 
          className="lg:hidden fixed inset-0 bg-black bg-opacity-50 z-40"
          onClick={onClose}
        />
      )}
      
      {/* Sidebar */}
      <div className={`
        w-72 bg-white border-r border-gray-200 h-full overflow-y-auto
        lg:relative lg:translate-x-0
        fixed z-50 transition-transform duration-300 ease-in-out
        ${isOpen ? 'translate-x-0' : '-translate-x-full'}
      `}>
        <div className="p-4">
          <h2 className="text-lg font-semibold text-gray-800 mb-4">FormulaEngine</h2>
          
          <nav className="space-y-1">
            {navigationData.map((item) => (
              <NavigationItemComponent
                key={item.id}
                item={item}
                level={0}
                currentPath={currentPath}
                onItemClick={onClose}
              />
            ))}
          </nav>
        </div>
        
        <div className="mt-auto p-4 border-t border-gray-200">
          <div className="text-xs text-gray-500">
            <div className="font-medium mb-1">FormulaEngine v1.0</div>
            <div>TypeScript Formula Engine</div>
            <div>with Excel Compatibility</div>
          </div>
        </div>
      </div>
    </>
  )
}
