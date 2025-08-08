import React, { useState, useMemo, useCallback } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { FormulaEngine, useSpreadsheet, useCell, useFormulaEngineEvents } from "../index";

export function EventsDemo() {
  // Create a FormulaEngine instance
  const engine = useMemo(() => {
    const eng = FormulaEngine.buildEmpty();
    const sheetName = eng.addSheet('Demo');
    const sheetId = eng.getSheetId(sheetName);
    
    // Set up some initial data
    const initialData = new Map<string, any>([
      ['A1', 10],
      ['B1', 20],
      ['C1', '=A1+B1'],
      ['A2', 5],
      ['B2', '=A1*2'],
      ['C2', '=B1+B2']
    ]);
    
    eng.setSheetContent(sheetId, initialData);
    return eng;
  }, []);

  // Use our React hooks
  const { spreadsheet, isLoading, error } = useSpreadsheet(engine, 'Demo');
  const a1Value = useCell(engine, 'Demo', 'A1');
  const c1Value = useCell(engine, 'Demo', 'C1');

  // Track events
  const [events, setEvents] = useState<any[]>([]);
  const [inputValue, setInputValue] = useState('');
  const [cellAddress, setCellAddress] = useState('A1');

  // Subscribe to events
  useFormulaEngineEvents(engine, {
    onCellChanged: useCallback((event: any) => {
      const addressStr = `${String.fromCharCode(65 + event.address.col)}${event.address.row + 1}`;
      setEvents(prev => [...prev.slice(-9), {
        type: 'cell-changed',
        timestamp: Date.now(),
        address: addressStr,
        oldValue: event.oldValue,
        newValue: event.newValue
      }]);
    }, []),
    
    onSheetAdded: useCallback((event: any) => {
      setEvents(prev => [...prev.slice(-9), {
        type: 'sheet-added',
        timestamp: Date.now(),
        sheetId: event.sheetId,
        sheetName: event.sheetName
      }]);
    }, []),
    
    onSheetRenamed: useCallback((event: any) => {
      setEvents(prev => [...prev.slice(-9), {
        type: 'sheet-renamed',
        timestamp: Date.now(),
        sheetId: event.sheetId,
        oldName: event.oldName,
        newName: event.newName
      }]);
    }, [])
  });

  const updateCell = () => {
    try {
      const sheetId = engine.getSheetId('Demo');
      const address = engine.simpleCellAddressFromString(cellAddress, sheetId);
      engine.setCellContent(address, inputValue);
      setInputValue('');
    } catch (error) {
      console.error('Error updating cell:', error);
    }
  };

  const addSheet = () => {
    const newSheetName = `Sheet${Date.now()}`;
    engine.addSheet(newSheetName);
  };

  const clearEvents = () => {
    setEvents([]);
  };

  if (isLoading) {
    return <div>Loading spreadsheet...</div>;
  }

  if (error) {
    return <div>Error: {error.message}</div>;
  }

  return (
    <div className="p-6 max-w-6xl mx-auto space-y-6">
      <div className="text-center">
        <h1 className="text-3xl font-bold mb-2">FormulaEngine Events & React Hooks Demo</h1>
        <p className="text-muted-foreground">
          Interactive demonstration of the events system and React integration
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Spreadsheet View */}
        <Card>
          <CardHeader>
            <CardTitle>Live Spreadsheet</CardTitle>
            <CardDescription>
              Using useSpreadsheet hook - updates automatically when values change
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-4 gap-2 mb-4">
              <div className="font-bold text-center"></div>
              <div className="font-bold text-center">A</div>
              <div className="font-bold text-center">B</div>
              <div className="font-bold text-center">C</div>
              
              <div className="font-bold text-center">1</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('A1') ?? ''}</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('B1') ?? ''}</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('C1') ?? ''}</div>
              
              <div className="font-bold text-center">2</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('A2') ?? ''}</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('B2') ?? ''}</div>
              <div className="text-center p-2 border rounded">{spreadsheet.get('C2') ?? ''}</div>
            </div>
            
            <div className="space-y-2">
              <div className="text-sm text-muted-foreground">
                Individual cell hooks:
              </div>
              <div className="text-sm">
                A1 (useCell): <span className="font-mono bg-gray-100 px-2 py-1 rounded">{a1Value.value}</span>
              </div>
              <div className="text-sm">
                C1 (useCell): <span className="font-mono bg-gray-100 px-2 py-1 rounded">{c1Value.value}</span>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Controls */}
        <Card>
          <CardHeader>
            <CardTitle>Controls</CardTitle>
            <CardDescription>
              Update cells and watch the events fire in real-time
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="space-y-2">
              <label className="text-sm font-medium">Cell Address:</label>
              <Input
                value={cellAddress}
                onChange={(e) => setCellAddress(e.target.value)}
                placeholder="e.g., A1, B2"
                className="font-mono"
              />
            </div>
            
            <div className="space-y-2">
              <label className="text-sm font-medium">Value or Formula:</label>
              <Input
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
                placeholder="e.g., 42, =A1*2, Hello"
                className="font-mono"
              />
            </div>
            
            <Button onClick={updateCell} className="w-full">
              Update Cell
            </Button>
            
            <div className="border-t pt-4">
              <Button onClick={addSheet} variant="outline" className="w-full mb-2">
                Add New Sheet
              </Button>
              <Button onClick={clearEvents} variant="outline" className="w-full">
                Clear Event Log
              </Button>
            </div>
            
            <div className="text-xs text-muted-foreground">
              Try these examples:
              <ul className="list-disc list-inside mt-1 space-y-1">
                <li>Set A1 to 100 and watch C1 update automatically</li>
                <li>Set B1 to a formula like =A1*3</li>
                <li>Add new sheets to see sheet-added events</li>
              </ul>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Events Log */}
      <Card>
        <CardHeader>
          <CardTitle>Event Log</CardTitle>
          <CardDescription>
            Real-time events from the FormulaEngine (showing last 10 events)
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="space-y-2 max-h-96 overflow-y-auto">
            {events.length === 0 ? (
              <div className="text-muted-foreground text-center py-8">
                No events yet. Try updating some cells!
              </div>
            ) : (
              events.slice().reverse().map((event, index) => (
                <div key={event.timestamp} className="flex items-center justify-between p-3 border rounded-lg">
                  <div className="flex-1">
                    <div className="font-mono text-sm">
                      <span className="font-semibold text-blue-600">{event.type}</span>
                      {event.type === 'cell-changed' && (
                        <span>
                          {' '}at {event.address}: {JSON.stringify(event.oldValue)} → {JSON.stringify(event.newValue)}
                        </span>
                      )}
                      {event.type === 'sheet-added' && (
                        <span>
                          {' '}"{event.sheetName}" (ID: {event.sheetId})
                        </span>
                      )}
                      {event.type === 'sheet-renamed' && (
                        <span>
                          {' '}"{event.oldName}" → "{event.newName}" (ID: {event.sheetId})
                        </span>
                      )}
                    </div>
                  </div>
                  <div className="text-xs text-muted-foreground">
                    {new Date(event.timestamp).toLocaleTimeString()}
                  </div>
                </div>
              ))
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}