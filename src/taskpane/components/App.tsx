import * as React from "react";
import { useState } from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";

const App = () => {
  const [selectedText, setSelectedText] = useState<string>("");
  const [wordCount, setWordCount] = useState<number>(0);
  const [status, setStatus] = useState<string>("");

  // Insert custom text
  const handleInsertText = async (text: string) => {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
      setStatus("Text inserted successfully!");
    }).catch((error) => {
      setStatus(`Error: ${error.message}`);
    });
  };

  // Get selected text
  const handleGetSelectedText = async () => {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      setSelectedText(body.text);
      setStatus("Selected text retrieved!");
    }).catch((error) => {
      setStatus(`Error: ${error.message}`);
    });
  };

  // Get selected text for TextInsertion component
  const getSelectedText = async (): Promise<string> => {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text || "";
    }).catch((error) => {
      console.error("Error getting selected text:", error);
      return "";
    });
  };

  // Count words in document
  const handleCountWords = async () => {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      const words = body.text
        .trim()
        .split(/\s+/)
        .filter((word) => word.length > 0).length;
      setWordCount(words);
      setStatus(`Document contains ${words} words!`);
    }).catch((error) => {
      setStatus(`Error: ${error.message}`);
    });
  };

  // Replace all selected text
  const handleReplaceText = async (newText: string) => {
    return Word.run(async (context) => {
      const range = context.document.body;
      range.clear();
      range.insertParagraph(newText, Word.InsertLocation.start);
      await context.sync();
      setStatus("Document replaced successfully!");
    }).catch((error) => {
      setStatus(`Error: ${error.message}`);
    });
  };

  // Apply formatting
  const handleApplyFormatting = async () => {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.font.bold = true;
      body.font.size = 14;
      await context.sync();
      setStatus("Formatting applied!");
    }).catch((error) => {
      setStatus(`Error: ${error.message}`);
    });
  };

  const heroItems: HeroListItem[] = [
    { icon: <span>📝</span>, primaryText: "Insert custom text into your document" },
    { icon: <span>📊</span>, primaryText: "Count words and analyze content" },
    { icon: <span>✏️</span>, primaryText: "Apply formatting to your text" },
    { icon: <span>📖</span>, primaryText: "View and manage document content" },
  ];

  return (
    <div className="ms-welcome" style={{ fontFamily: "Segoe UI, sans-serif" }}>
      <Header title="Word Add-in" logo="assets/getaway_logo_black.svg" message="Document Tools" />
      {/* <HeroList message="Available Features" items={heroItems} /> */}

      <div style={{ padding: "20px", maxWidth: "600px", margin: "0 auto" }}>
        <TextInsertion insertText={handleInsertText} getSelectedText={getSelectedText} />

        <div style={{ marginTop: "30px", display: "flex", flexDirection: "column", gap: "10px" }}>
          <button
            onClick={handleGetSelectedText}
            style={{
              padding: "10px 20px",
              backgroundColor: "#0078d4",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
            }}
          >
            Get Document Text
          </button>

          <button
            onClick={handleCountWords}
            style={{
              padding: "10px 20px",
              backgroundColor: "#107c10",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
            }}
          >
            Count Words
          </button>

          <button
            onClick={handleApplyFormatting}
            style={{
              padding: "10px 20px",
              backgroundColor: "#d83b01",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
            }}
          >
            Apply Bold & Larger Font
          </button>

          {status && (
            <div
              style={{
                marginTop: "15px",
                padding: "10px",
                backgroundColor: "#f0f0f0",
                borderRadius: "4px",
                color: "#333",
              }}
            >
              {status}
            </div>
          )}

          {selectedText && (
            <div
              style={{
                marginTop: "15px",
                padding: "10px",
                backgroundColor: "#e8f4f8",
                borderRadius: "4px",
                color: "#333",
              }}
            >
              <strong>Document Text:</strong>
              <p>{selectedText.substring(0, 200)}...</p>
            </div>
          )}

          {wordCount > 0 && (
            <div
              style={{
                marginTop: "15px",
                padding: "10px",
                backgroundColor: "#e8f8e8",
                borderRadius: "4px",
                color: "#333",
              }}
            >
              <strong>Word Count: {wordCount} words</strong>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default App;
