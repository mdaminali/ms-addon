import * as React from "react";

const App = () => {
  const click = async () => {
    return Word.run(async (context) => {
      const body = context.document.body;

      body.insertParagraph("Hello World from React!", Word.InsertLocation.end);

      await context.sync();
    });
  };

  return (
    <div className="ms-welcome">
      <h1>Welcome to My Add-in</h1>
      <p>Click the button to insert text.</p>

      {/* Standard HTML Button */}
      <button onClick={click} style={{ padding: "10px 20px", cursor: "pointer" }}>
        Insert Hello World
      </button>
    </div>
  );
};

export default App;
