import { createFileRoute, Link } from '@tanstack/react-router'

export const Route = createFileRoute('/')({
  component: Index,
})

function Index() {
  return (
    <div className="max-w-6xl mx-auto">
      <div className="text-center mb-12">
        <h1 className="text-4xl font-bold mb-6 text-gray-800">Welcome to FormulaEngine</h1>
        <p className="text-lg text-gray-600 max-w-3xl mx-auto">
          A powerful TypeScript-based spreadsheet formula evaluation library with sparse-aware evaluation, 
          Excel compatibility, and advanced features like array formulas and cross-sheet references.
        </p>
      </div>
      
      <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-4 gap-6 mb-12">
        <Link 
          to="/events" 
          className="group p-6 bg-white border rounded-xl hover:shadow-lg transition-all duration-200 hover:border-blue-200"
        >
          <div className="flex items-center mb-4">
            <div className="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center mr-3">
              <span className="text-blue-600 font-semibold">🔗</span>
            </div>
            <h3 className="text-lg font-semibold text-gray-800 group-hover:text-blue-600">Events & Hooks</h3>
          </div>
          <p className="text-gray-600 leading-relaxed text-sm">
            Explore React integration with custom hooks and event system for reactive spreadsheet updates.
          </p>
        </Link>
        
        <Link 
          to="/spreadsheet" 
          className="group p-6 bg-white border rounded-xl hover:shadow-lg transition-all duration-200 hover:border-green-200"
        >
          <div className="flex items-center mb-4">
            <div className="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center mr-3">
              <span className="text-green-600 font-semibold">📊</span>
            </div>
            <h3 className="text-lg font-semibold text-gray-800 group-hover:text-green-600">Full Spreadsheet</h3>
          </div>
          <p className="text-gray-600 leading-relaxed text-sm">
            Experience a complete spreadsheet interface with formula evaluation, cell editing, and real-time calculations.
          </p>
        </Link>
        
        <Link 
          to="/multisheet" 
          className="group p-6 bg-white border rounded-xl hover:shadow-lg transition-all duration-200 hover:border-purple-200"
        >
          <div className="flex items-center mb-4">
            <div className="w-10 h-10 bg-purple-100 rounded-lg flex items-center justify-center mr-3">
              <span className="text-purple-600 font-semibold">📚</span>
            </div>
            <h3 className="text-lg font-semibold text-gray-800 group-hover:text-purple-600">Multi-Sheet</h3>
          </div>
          <p className="text-gray-600 leading-relaxed text-sm">
            Test cross-sheet references and multi-sheet functionality with our advanced formula engine.
          </p>
        </Link>
        
        <Link 
          to="/excel" 
          className="group p-6 bg-white border rounded-xl hover:shadow-lg transition-all duration-200 hover:border-orange-200"
        >
          <div className="flex items-center mb-4">
            <div className="w-10 h-10 bg-orange-100 rounded-lg flex items-center justify-center mr-3">
              <span className="text-orange-600 font-semibold">📋</span>
            </div>
            <h3 className="text-lg font-semibold text-gray-800 group-hover:text-orange-600">Excel Clone</h3>
          </div>
          <p className="text-gray-600 leading-relaxed text-sm">
            Full Excel-like experience with sheet tabs, add/rename/delete sheets, and familiar Excel UI.
          </p>
        </Link>
      </div>
      
      <div className="bg-white border rounded-xl p-8">
        <h2 className="text-2xl font-semibold mb-6 text-gray-800">Key Features</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div>
            <h3 className="font-semibold text-gray-800 mb-2">🚀 Performance</h3>
            <p className="text-gray-600 text-sm">Sparse-aware evaluation processes only populated cells for optimal memory usage and speed.</p>
          </div>
          <div>
            <h3 className="font-semibold text-gray-800 mb-2">🎯 Excel Compatibility</h3>
            <p className="text-gray-600 text-sm">Comprehensive function library matching Excel behavior for seamless migration.</p>
          </div>
          <div>
            <h3 className="font-semibold text-gray-800 mb-2">📐 Array Formulas</h3>
            <p className="text-gray-600 text-sm">NumPy-style broadcasting with automatic spilling for advanced calculations.</p>
          </div>
          <div>
            <h3 className="font-semibold text-gray-800 mb-2">🔧 TypeScript</h3>
            <p className="text-gray-600 text-sm">Full type safety with advanced TypeScript features and IntelliSense support.</p>
          </div>
        </div>
      </div>
    </div>
  )
}