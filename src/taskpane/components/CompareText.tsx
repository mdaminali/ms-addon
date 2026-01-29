import * as React from "react";
import { useState } from "react";
import { diffWords } from "diff";

interface CompareTextProps {
  setStatus: (status: string) => void;
}

interface DiffPart {
  value: string;
  added?: boolean;
  removed?: boolean;
}

const CompareText: React.FC<CompareTextProps> = (props: CompareTextProps) => {
  const [mainText, setMainText] = useState<string>(
    "Bangladesh is a land of rivers, greenery, and vibrant culture."
  );
  const [changeText, setChangeText] = useState<string>(
    "Bangladesh is a land of flowing rivers, lush landscapes, and rich heritage."
  );
  const [comparisonResult, setComparisonResult] = useState<DiffPart[]>([]);
  const [showComparison, setShowComparison] = useState<boolean>(false);

  const handleCompare = () => {
    const differences = diffWords(mainText, changeText);
    setComparisonResult(differences);
    setShowComparison(true);
  };

  const handleInsertChange = async () => {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(changeText, Word.InsertLocation.end);
      await context.sync();
      props.setStatus("New text inserted successfully!");
    }).catch((error) => {
      props.setStatus(`Error: ${error.message}`);
    });
  };

  return (
    <div
      style={{
        marginTop: "30px",
        marginBottom: "20px",
        border: "1px solid #ccc",
        padding: "15px",
        borderRadius: "8px",
      }}
    >
      <h3>Compare Text</h3>

      <label style={{ display: "block", marginBottom: "5px", fontWeight: "bold" }}>
        Original Text (Main):
      </label>
      <textarea
        value={mainText}
        onChange={(e) => setMainText(e.target.value)}
        style={{ width: "100%", height: "60px", marginBottom: "10px", padding: "5px" }}
      />

      <label style={{ display: "block", marginBottom: "5px", fontWeight: "bold" }}>
        New Text (Changed):
      </label>
      <textarea
        value={changeText}
        onChange={(e) => setChangeText(e.target.value)}
        style={{ width: "100%", height: "60px", marginBottom: "15px", padding: "5px" }}
      />

      <div style={{ display: "flex", gap: "10px", marginBottom: "15px" }}>
        <button
          onClick={handleCompare}
          style={{
            flex: 1,
            padding: "12px",
            backgroundColor: "#4b0082",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
          }}
        >
          Compare
        </button>

        <button
          onClick={handleInsertChange}
          style={{
            flex: 1,
            padding: "12px",
            backgroundColor: "#107c10",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
          }}
        >
          Insert Change
        </button>
      </div>

      {/* COMPARISON RESULT */}
      {showComparison && comparisonResult.length > 0 && (
        <div
          style={{
            marginTop: "15px",
            padding: "15px",
            backgroundColor: "#f5f5f5",
            borderRadius: "4px",
            border: "1px solid #ddd",
          }}
        >
          <h4 style={{ marginTop: 0 }}>Comparison Result:</h4>
          <div
            style={{
              lineHeight: "1.8",
              fontSize: "14px",
              wordBreak: "break-word",
            }}
          >
            {comparisonResult.map((part, index) => {
              let style: React.CSSProperties = { display: "inline" };

              if (part.added) {
                style.color = "green";
                style.fontWeight = "bold";
                style.textDecoration = "underline";
              } else if (part.removed) {
                style.color = "#D00000";
                style.textDecoration = "line-through";
              }

              return (
                <span key={index} style={style}>
                  {part.value}
                </span>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
};

export default CompareText;
