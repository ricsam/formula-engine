import "./index.css";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { EventsDemo } from "./EventsDemo";
import { FullSpreadsheetDemo } from "./FullSpreadsheetDemo";
import { MultiSheetDemo } from "./MultiSheetDemo";

export function App() {
  const [currentView, setCurrentView] = useState<'spreadsheet' | 'events' | 'multisheet'>('events');

  return (
    <div className="container mx-auto p-4 relative z-10 flex flex-col gap-4 h-screen">
      {/* Tab Navigation */}
      <div className="flex justify-center gap-2 mb-4">
        <Button 
          variant={currentView === 'events' ? 'default' : 'outline'}
          onClick={() => setCurrentView('events')}
        >
          Events & Hooks Demo
        </Button>
        <Button 
          variant={currentView === 'spreadsheet' ? 'default' : 'outline'}
          onClick={() => setCurrentView('spreadsheet')}
        >
          Full Spreadsheet
        </Button>
        <Button 
          variant={currentView === 'multisheet' ? 'default' : 'outline'}
          onClick={() => setCurrentView('multisheet')}
        >
          Multi-Sheet Demo
        </Button>
      </div>

      {/* Content */}
      {currentView === 'events' ? (
        <EventsDemo />
      ) : currentView === 'spreadsheet' ? (
        <FullSpreadsheetDemo />
      ) : (
        <MultiSheetDemo />
      )}
    </div>
  );
}

export default App;
