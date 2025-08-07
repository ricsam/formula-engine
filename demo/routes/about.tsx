import { createFileRoute } from '@tanstack/react-router'

export const Route = createFileRoute('/about')({
  component: About,
})

function About() {
  return (
    <div className="max-w-4xl mx-auto p-6">
      <h1 className="text-3xl font-bold mb-6">About FormulaEngine</h1>
      
      <div className="prose prose-lg max-w-none">
        <p className="text-lg text-gray-700 mb-6">
          FormulaEngine is a high-performance, TypeScript-based spreadsheet formula evaluation library 
          designed for modern web applications requiring sophisticated computational capabilities.
        </p>

        <h2 className="text-2xl font-semibold mb-4">Key Features</h2>
        <ul className="list-disc pl-6 mb-6 space-y-2">
          <li><strong>Sparse-Aware Evaluation:</strong> Only defined cells consume memory and processing resources</li>
          <li><strong>Excel Compatibility:</strong> Comprehensive function library matching Excel behavior</li>
          <li><strong>TypeScript First:</strong> Full type safety with advanced TypeScript features</li>
          <li><strong>Array Formulas:</strong> NumPy-style broadcasting with automatic spilling</li>
          <li><strong>Cross-Sheet References:</strong> Support for formulas spanning multiple sheets</li>
          <li><strong>React Integration:</strong> Custom hooks for reactive spreadsheet components</li>
          <li><strong>Dependency Tracking:</strong> Intelligent recalculation with circular reference detection</li>
        </ul>

        <h2 className="text-2xl font-semibold mb-4">Function Categories</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Mathematical Functions</h3>
            <p className="text-sm text-gray-600">Basic arithmetic, advanced math, and statistical functions</p>
          </div>
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Logical Functions</h3>
            <p className="text-sm text-gray-600">Conditional logic, boolean operations, and comparisons</p>
          </div>
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Text Functions</h3>
            <p className="text-sm text-gray-600">String manipulation, formatting, and text processing</p>
          </div>
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Lookup Functions</h3>
            <p className="text-sm text-gray-600">VLOOKUP, INDEX, MATCH, and advanced lookup capabilities</p>
          </div>
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Array Functions</h3>
            <p className="text-sm text-gray-600">FILTER, array manipulation, and dynamic arrays</p>
          </div>
          <div className="p-4 border rounded">
            <h3 className="font-semibold mb-2">Info Functions</h3>
            <p className="text-sm text-gray-600">Type checking, error detection, and metadata functions</p>
          </div>
        </div>

        <h2 className="text-2xl font-semibold mb-4">Technical Highlights</h2>
        <ul className="list-disc pl-6 mb-6 space-y-2">
          <li>Three-phase processing: Parsing → Dependency Analysis → Evaluation</li>
          <li>Compressed Sparse Row (CSR) format for optimal memory usage</li>
          <li>Multi-level caching with intelligent invalidation</li>
          <li>Incremental recalculation for performance</li>
          <li>AST reuse for common formula patterns</li>
          <li>Comprehensive error handling with Excel-compatible error types</li>
        </ul>

        <div className="bg-blue-50 p-4 rounded-lg">
          <p className="text-blue-800">
            <strong>Performance:</strong> Designed to handle spreadsheets with 100,000+ formulas 
            with sub-second recalculation times while using 10x less memory than traditional dense storage.
          </p>
        </div>
      </div>
    </div>
  )
}