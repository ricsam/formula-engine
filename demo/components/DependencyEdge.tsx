import React, { type FC, useState, useRef, useEffect } from "react";
import {
  getBezierPath,
  EdgeLabelRenderer,
  BaseEdge,
  type EdgeProps,
  type Edge,
} from "@xyflow/react";
import { Card } from "@/components/ui/card";
import { enhancedFormulaForDisplay } from "../lib/dependency-analyzer";

interface DependencyEdgeData extends Record<string, unknown> {
  formulas: Array<{ formula: string; cellAddress: string }>;
  cellCount: number;
}

const DependencyEdge: FC<EdgeProps<Edge<DependencyEdgeData>>> = ({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
  sourcePosition,
  targetPosition,
  data,
  style,
  markerEnd,
}) => {
  const [showDetails, setShowDetails] = useState(false);

  const [edgePath, labelX, labelY] = getBezierPath({
    sourceX,
    sourceY,
    sourcePosition,
    targetX,
    targetY,
    targetPosition,
  });

  return (
    <>
      <BaseEdge id={id} path={edgePath} style={style} markerEnd={markerEnd} />
      <EdgeLabelRenderer>
        <div
          style={{
            position: "absolute",
            transform: `translate(-50%, -50%) translate(${labelX}px,${labelY}px)`,
            pointerEvents: "all",
          }}
          className="nodrag nopan"
        >
          <button
            onClick={() => setShowDetails(!showDetails)}
            className="bg-blue-100 hover:bg-blue-200 text-blue-800 text-xs px-2 py-1 rounded border border-blue-300 transition-colors cursor-pointer"
          >
            {data?.cellCount || 0} formulas
          </button>

          {showDetails && data?.formulas && (
            <div
              style={{
                position: "fixed",
                top: "50%",
                left: "50%",
                transform: "translate(-50%, 16px)",
                zIndex: 1000,
                width: "320px",
                maxWidth: "1200px",
              }}
            >
              <Card className="p-2 w-full shadow-lg bg-white border">
                <div
                  className="space-y-1 text-[8px] font-mono max-h-[70vh] overflow-y-auto pr-2"
                  style={{ scrollbarWidth: "thin" }}
                >
                  {data.formulas.map((formulaData, idx) => (
                    <div
                      key={`${formulaData.cellAddress}-${idx}`}
                      className="bg-gray-50 p-1.5 rounded break-all hover:bg-gray-100 transition-colors"
                    >
                      <div className="flex justify-start gap-2">
                        <span className="text-gray-500">{formulaData.cellAddress}</span>
                        <div className="text-gray-900 flex-1 whitespace-pre-wrap">
                          {enhancedFormulaForDisplay(
                            formulaData.formula,
                            formulaData.cellAddress
                          )}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </Card>
            </div>
          )}
        </div>
      </EdgeLabelRenderer>
    </>
  );
};

export default DependencyEdge;
