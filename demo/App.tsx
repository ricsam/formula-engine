import "./index.css";
import { EventsDemo } from "./EventsDemo";
import { FullSpreadsheetDemo } from "./FullSpreadsheetDemo";
import { MultiSheetDemo } from "./MultiSheetDemo";

/**
 * Legacy App component - now replaced by TanStack Router
 * This component is kept for reference but is no longer used in the routing system
 * Individual demo components are now accessed via their respective routes:
 * - /events -> EventsDemo
 * - /spreadsheet -> FullSpreadsheetDemo  
 * - /multisheet -> MultiSheetDemo
 */
export function App() {
  return (
    <div className="container mx-auto p-4 relative z-10 flex flex-col gap-4 h-screen">
      <div className="text-center p-8">
        <h1 className="text-2xl font-bold mb-4">FormulaEngine Demo</h1>
        <p className="text-gray-600">
          This component is now replaced by TanStack Router. 
          Please use the navigation above to access the different demos.
        </p>
      </div>
    </div>
  );
}

export default App;

// Re-export demo components for easy access
export { EventsDemo, FullSpreadsheetDemo, MultiSheetDemo };
