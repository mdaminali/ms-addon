import * as React from "react";
import { useState, useEffect } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
  getSelectedText?: () => Promise<string>;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    margin: "10px",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");
  const [selectedText, setSelectedText] = useState<string>("");

  // useEffect(() => {
  //   loadSelectedText();
  // }, [props]);

  const loadSelectedText = async () => {
    if (props.getSelectedText) {
      try {
        const selectedText = await props.getSelectedText();
        if (selectedText) {
          setSelectedText(selectedText);
        }
      } catch (error) {
        console.error("Error loading selected text:", error);
      }
    }
  };

  const handleTextInsertion = async () => {
    await props.insertText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field
        className={styles.textAreaField}
        size="large"
        label="Enter text to be inserted into the document."
      >
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert text
      </Button>

      <Button
        style={{ marginTop: "10px" }}
        appearance="secondary"
        disabled={false}
        size="large"
        onClick={loadSelectedText}
      >
        Get selected text
      </Button>
      <p>{selectedText}</p>
    </div>
  );
};

export default TextInsertion;
