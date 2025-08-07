import React, { useMemo, useCallback, useState, useEffect } from "react";
import {
  ReactFlow,
  useNodesState,
  useEdgesState,
  addEdge,
  ConnectionLineType,
  Panel,
  Background,
  Controls,
  MiniMap,
  Handle,
  Position,
  BackgroundVariant,
  MarkerType,
  type EdgeTypes,
  type Edge,
  type Node,
  type Connection,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { createEngineWithMultiSheetData } from "./lib/multisheet-data";
import { analyzeDependencies } from "./lib/dependency-analyzer";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import DependencyEdge from "./components/DependencyEdge";

// Custom node component for sheet representation
const SheetNode = ({ data }: { data: any }) => {
  const handleDownloadCSV = () => {
    // TODO: Implement CSV download functionality
    console.log(`Downloading CSV for ${data.name}`);
    alert(`Download CSV for ${data.name} - Feature coming soon!`);
  };

  const handleUploadCSV = () => {
    // TODO: Implement CSV upload functionality
    console.log(`Uploading CSV for ${data.name}`);
    alert(`Upload CSV for ${data.name} - Feature coming soon!`);
  };

  return (
    <div className="px-6 py-4 shadow-lg rounded-lg bg-white border-2 border-gray-200 hover:border-blue-300 transition-colors relative">
      {/* Input handle (top) - for edges coming to this node */}
      <Handle
        type="target"
        position={Position.Top}
        id="top"
        className="!w-3 !h-3 !bg-blue-500 !border-2 !border-white"
      />
      
      {/* Output handle (bottom) - for edges going from this node */}
      <Handle
        type="source"
        position={Position.Bottom}
        id="bottom"
        className="!w-3 !h-3 !bg-blue-500 !border-2 !border-white"
      />
      
      <div className="flex flex-col items-center space-y-3">
        <div className="text-2xl">{data.emoji}</div>
        <div className="font-bold text-lg text-gray-800">{data.name}</div>
        <div className="text-xs text-gray-500">Sheet</div>
        
        {/* Action buttons based on dependency relationships */}
        <div className="flex flex-col gap-1 w-full">
          {data.hasIncomingEdges && (
            <Button
              onClick={handleDownloadCSV}
              size="sm"
              variant="outline"
              className="text-xs h-6 px-2"
            >
              📥 Download CSV
            </Button>
          )}
          {data.hasOutgoingEdges && (
            <Button
              onClick={handleUploadCSV}
              size="sm"
              variant="outline"
              className="text-xs h-6 px-2"
            >
              📤 Upload CSV
            </Button>
          )}
        </div>
      </div>
    </div>
  );
};

const initialNodes: Node[] = [
  {
    id: "Products",
    type: "sheetNode",
    position: { x: 100, y: 100 },
    data: { name: "Products", emoji: "📦" },
  },
  {
    id: "Sales",
    type: "sheetNode",
    position: { x: 400, y: 100 },
    data: { name: "Sales", emoji: "💰" },
  },
  {
    id: "Dashboard",
    type: "sheetNode",
    position: { x: 250, y: 300 },
    data: { name: "Dashboard", emoji: "📊" },
  },
];

const initialEdges: Edge[] = [];

// Define nodeTypes and edgeTypes outside component to prevent recreation on every render
const nodeTypes = {
  sheetNode: SheetNode,
};

const edgeTypes: EdgeTypes = {
  dependency: DependencyEdge,
};

export function DependencyFlowDemo() {
  const { engine, sheets } = useMemo(createEngineWithMultiSheetData, []);
  const [edges, setEdges, onEdgesChange] = useEdgesState(initialEdges);
  const [showDependencies, setShowDependencies] = useState(true);

  const dependencyGraph = useMemo(() => {
    const graph = analyzeDependencies(engine, sheets);
    console.log("Dependency Graph:", graph);
    console.log("Sheets:", sheets);
    return graph;
  }, [engine, sheets]);

  // Create nodes with dependency information
  const nodesWithDependencies = useMemo(() => {
    const incomingNodes = new Set(dependencyGraph.edges.map(edge => edge.target));
    const outgoingNodes = new Set(dependencyGraph.edges.map(edge => edge.source));
    
    return initialNodes.map(node => ({
      ...node,
      data: {
        ...node.data,
        hasIncomingEdges: incomingNodes.has(node.id),
        hasOutgoingEdges: outgoingNodes.has(node.id),
      }
    }));
  }, [dependencyGraph.edges]);

  const [nodes, setNodes, onNodesChange] = useNodesState(nodesWithDependencies);

  // Update nodes when dependencies change
  useEffect(() => {
    setNodes(nodesWithDependencies);
  }, [nodesWithDependencies, setNodes]);

  const onConnect = useCallback(
    (params: Edge | Connection) => setEdges((eds) => addEdge(params, eds)),
    [setEdges]
  );

  // Show dependencies by default on mount
  useEffect(() => {
    if (dependencyGraph.edges.length > 0) {
      const dependencyEdges: Edge[] = dependencyGraph.edges.map((edge) => ({
        id: edge.id,
        source: edge.source,
        target: edge.target,
        sourceHandle: "bottom",
        targetHandle: "top",
        type: "dependency",
        animated: true,
        style: {
          stroke: "#3b82f6",
          strokeWidth: 3,
        },
        markerEnd: {
          type: MarkerType.ArrowClosed,
          color: "#3b82f6",
        },
        data: {
          formulas: edge.formulas,
          cellCount: edge.cellCount,
        },
      }));

      setEdges(dependencyEdges);
    }
  }, [dependencyGraph.edges, setEdges]);

  const toggleDependencies = useCallback(() => {
    if (!showDependencies) {
      // Show dependencies - create edges with custom labels
      const dependencyEdges: Edge[] = dependencyGraph.edges.map((edge) => ({
        id: edge.id,
        source: edge.source,
        target: edge.target,
        sourceHandle: "bottom",
        targetHandle: "top",
        type: "dependency",
        animated: true,
        style: {
          stroke: "#3b82f6",
          strokeWidth: 3,
        },
        markerEnd: {
          type: MarkerType.ArrowClosed,
          color: "#3b82f6",
        },
        data: {
          formulas: edge.formulas,
          cellCount: edge.cellCount,
        },
      }));

      setEdges(dependencyEdges);
      setShowDependencies(true);
    } else {
      // Hide dependencies
      setEdges([]);
      setShowDependencies(false);
    }
  }, [showDependencies, dependencyGraph.edges, setEdges]);

  const layoutNodes = useCallback(() => {
    // Simple circular layout for the three sheets
    const centerX = 400;
    const centerY = 250;
    const radius = 150;

    const updatedNodes = nodes.map((node, index) => {
      const angle = (index * 2 * Math.PI) / nodes.length - Math.PI / 2; // Start from top
      return {
        ...node,
        position: {
          x: centerX + radius * Math.cos(angle) - 75, // Center the node (assuming 150px width)
          y: centerY + radius * Math.sin(angle) - 50, // Center the node (assuming 100px height)
        },
      };
    });

    setNodes(updatedNodes);
  }, [nodes, setNodes]);

  return (
    <div className="w-full h-full flex flex-col">
      {/* Header */}
      <div className="bg-white border-b p-4 flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold">Sheet Dependency Graph</h1>
          <p className="text-gray-600 text-sm">
            Visualize cross-sheet formula dependencies in your spreadsheet
          </p>
        </div>
        <div className="flex gap-2">
          <Button onClick={layoutNodes} variant="outline" size="sm">
            Auto Layout
          </Button>
          <Button
            onClick={toggleDependencies}
            variant={showDependencies ? "default" : "outline"}
            size="sm"
          >
            {showDependencies ? "Hide Dependencies" : "Show Dependencies"}
          </Button>
        </div>
      </div>

      {/* React Flow */}
      <div className="flex-1 bg-gray-50">
        <ReactFlow
          nodes={nodes}
          edges={edges}
          onNodesChange={onNodesChange}
          onEdgesChange={onEdgesChange}
          onConnect={onConnect}
          nodeTypes={nodeTypes}
          edgeTypes={edgeTypes}
          connectionLineType={ConnectionLineType.SmoothStep}
          fitView
          fitViewOptions={{ padding: 0.2 }}
        >
          <Controls />
          <MiniMap
            nodeColor={(node) => {
              switch (node.id) {
                case "Products":
                  return "#10b981";
                case "Sales":
                  return "#f59e0b";
                case "Dashboard":
                  return "#3b82f6";
                default:
                  return "#6b7280";
              }
            }}
          />
          <Background variant={BackgroundVariant.Dots} gap={12} size={1} />

          <Panel
            position="bottom-left"
            className="bg-white p-3 rounded shadow-lg border"
          >
            <div className="text-sm space-y-1">
              <div className="flex items-center gap-2">
                <div className="w-3 h-3 rounded bg-green-500"></div>
                <span className="text-xs">📦 Products (Source Data)</span>
              </div>
              <div className="flex items-center gap-2">
                <div className="w-3 h-3 rounded bg-amber-500"></div>
                <span className="text-xs">💰 Sales (Transactions)</span>
              </div>
              <div className="flex items-center gap-2">
                <div className="w-3 h-3 rounded bg-blue-500"></div>
                <span className="text-xs">📊 Dashboard (Analytics)</span>
              </div>
              {showDependencies && (
                <div className="flex items-center gap-2 mt-2 pt-2 border-t">
                  <div className="w-3 h-0.5 bg-blue-500"></div>
                  <span className="text-xs">Data Flow Direction</span>
                </div>
              )}
            </div>
          </Panel>
        </ReactFlow>
      </div>
    </div>
  );
}
