import { useRef, useState } from "react";
import "./App.css";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
} from "docx";

interface ProcessedFile {
  name: string;
  contentType: string;
  size: number;
  base64: string;
}

interface LogicAppPayload {
  prompt: string;
  files: ProcessedFile[];
}

function App() {
  const fileInputRef =
    useRef<HTMLInputElement>(null);

  const [prompt, setPrompt] =
    useState("");

  const [selectedFiles, setSelectedFiles] =
    useState<File[]>([]);

  const [processedFiles, setProcessedFiles] =
    useState<ProcessedFile[]>([]);

  /*
   * Editable Logic App response.
   */
  const [response, setResponse] =
    useState("");

  const [isLoading, setIsLoading] =
    useState(false);

  const fileToBase64 = (
    file: File
  ): Promise<string> => {
    return new Promise(
      (resolve, reject) => {
        const reader =
          new FileReader();

        reader.onload = () => {
          const result =
            reader.result;

          if (
            typeof result !==
            "string"
          ) {
            reject(
              new Error(
                `Unable to read file: ${file.name}`
              )
            );
            return;
          }

          const base64 =
            result.split(",")[1];

          if (!base64) {
            reject(
              new Error(
                `Unable to extract Base64: ${file.name}`
              )
            );
            return;
          }

          resolve(base64);
        };

        reader.onerror = () => {
          reject(
            reader.error ??
              new Error(
                `Unable to read file: ${file.name}`
              )
          );
        };

        reader.readAsDataURL(file);
      }
    );
  };

  const handleUploadClick =
    () => {
      fileInputRef.current?.click();
    };

  const handlePromptChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>
  ) => {
    setPrompt(
      event.target.value
    );
  };

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = Array.from(
      event.target.files ?? []
    );

    if (files.length === 0) {
      return;
    }

    try {
      const newProcessedFiles: ProcessedFile[] =
        [];

      for (const file of files) {
        const base64 =
          await fileToBase64(file);

        newProcessedFiles.push({
          name: file.name,
          contentType:
            file.type,
          size: file.size,
          base64,
        });
      }

      setSelectedFiles(
        (previousFiles) => [
          ...previousFiles,
          ...files,
        ]
      );

      setProcessedFiles(
        (previousFiles) => [
          ...previousFiles,
          ...newProcessedFiles,
        ]
      );
    } catch (error) {
      console.error(
        "Error processing file(s):",
        error
      );
    }

    /*
     * Allows the same file to be
     * selected again later.
     */
    event.target.value = "";
  };

  const handleRemoveFile = (
    indexToRemove: number
  ) => {
    setSelectedFiles(
      (previousFiles) =>
        previousFiles.filter(
          (_, index) =>
            index !==
            indexToRemove
        )
    );

    setProcessedFiles(
      (previousFiles) =>
        previousFiles.filter(
          (_, index) =>
            index !==
            indexToRemove
        )
    );
  };

  const buildPayload =
    (): LogicAppPayload => {
      return {
        prompt: prompt.trim(),
        files: processedFiles,
      };
    };

  /*
   * Call Logic App.
   */
  const handleGetResponse =
    async () => {
      if (!prompt.trim()) {
        alert(
          "Please enter an instruction or query."
        );
        return;
      }

      if (
        processedFiles.length ===
        0
      ) {
        alert(
          "Please upload at least one file."
        );
        return;
      }

      const logicAppUrl =
        import.meta.env
          .VITE_LOGIC_APP_URL;

      if (!logicAppUrl) {
        alert(
          "Logic App URL is not configured."
        );
        return;
      }

      setIsLoading(true);
      setResponse("");

      try {
        const payload =
          buildPayload();

        const result =
          await fetch(
            logicAppUrl,
            {
              method: "POST",

              headers: {
                "Content-Type":
                  "application/json",
              },

              body: JSON.stringify(
                payload
              ),
            }
          );

        const responseText =
          await result.text();

        if (!result.ok) {
          throw new Error(
            `Logic App returned HTTP ${result.status}: ${responseText}`
          );
        }

        /*
         * Try JSON first.
         */
        try {
          const jsonResponse =
            JSON.parse(
              responseText
            );

          /*
           * If Logic App returns an
           * object containing "response",
           * use that response.
           */
          if (
            jsonResponse &&
            typeof jsonResponse ===
              "object" &&
            "response" in
              jsonResponse
          ) {
            setResponse(
              String(
                jsonResponse.response
              )
            );
          } else {
            /*
             * Otherwise display the
             * formatted JSON.
             */
            setResponse(
              JSON.stringify(
                jsonResponse,
                null,
                2
              )
            );
          }
        } catch {
          /*
           * Logic App returned
           * plain text.
           */
          setResponse(
            responseText
          );
        }
      } catch (error) {
        console.error(
          "Logic App request failed:",
          error
        );

        setResponse(
          error instanceof Error
            ? error.message
            : "An unexpected error occurred."
        );
      } finally {
        setIsLoading(false);
      }
    };

  /*
   * Download the edited response
   * as a DOCX file.
   */
  const handleDownloadDocx = async () => {
  if (!response.trim()) {
    alert("There is no response to download.");
    return;
  }

  try {
    const lines = response.split(/\r?\n/);

    const paragraphs = lines.map(
      (line) =>
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [
            new TextRun({
              text: line,
              size: 22,
            }),
          ],
        })
    );

    const docxDocument = new Document({
      sections: [
        {
          properties: {},
          children: paragraphs,
        },
      ],
    });

    const blob = await Packer.toBlob(
      docxDocument
    );

    const downloadUrl =
      URL.createObjectURL(blob);

    const link =
      window.document.createElement("a");

    link.href = downloadUrl;
    link.download = "HLID_Response.docx";

    window.document.body.appendChild(link);

    link.click();

    window.document.body.removeChild(link);

    URL.revokeObjectURL(downloadUrl);

  } catch (error) {
    console.error(
      "DOCX generation failed:",
      error
    );

    alert(
      "Unable to generate the Word document."
    );
  }
};

  return (
    <div className="app-container">

      {/* Header */}

      <header className="app-header">

        <h1>
          Welcome at HLID Assistant
        </h1>

        <div className="user-name">
          User Name
        </div>

      </header>

      {/* Main Content */}

      <main className="main-content">

        {/* Prompt */}

        <textarea
          className="query-box"
          placeholder="Enter your instruction or query here..."
          value={prompt}
          onChange={
            handlePromptChange
          }
        />

        {/* Hidden file input */}

        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept="
            .c,
            .cpp,
            .cs,
            .css,
            .doc,
            .docx,
            .go,
            .html,
            .java,
            .js,
            .json,
            .md,
            .pdf,
            .php,
            .pptx,
            .py,
            .rb,
            .sh,
            .tex,
            .ts,
            .txt,
            .jpg,
            .jpeg,
            .png,
            .bmp,
            .tiff,
            .heif
          "
          style={{
            display: "none",
          }}
          onChange={
            handleFileChange
          }
        />

        {/* Selected Files */}

        {selectedFiles.length >
          0 && (
          <div className="selected-files">

            <div className="selected-files-title">
              Selected File(s)
            </div>

            {selectedFiles.map(
              (
                file,
                index
              ) => (
                <div
                  className="selected-file"
                  key={`${file.name}-${index}`}
                >

                  <span className="selected-file-name">
                    {file.name}
                  </span>

                  <button
                    type="button"
                    className="remove-file-button"
                    onClick={() =>
                      handleRemoveFile(
                        index
                      )
                    }
                    title={`Remove ${file.name}`}
                    aria-label={`Remove ${file.name}`}
                  >
                    ×
                  </button>

                </div>
              )
            )}

          </div>
        )}

        {/* Buttons */}

        <div className="button-container">

          <button
            type="button"
            className="primary-button upload-button"
            onClick={
              handleUploadClick
            }
            disabled={isLoading}
            title="Allowable file formats: .c, .cpp, .cs, .css, .doc, .docx, .go, .html, .java, .js, .json, .md, .pdf, .php, .pptx, .py, .rb, .sh, .tex, .ts, .txt, .jpg, .jpeg, .png, .bmp, .tiff, .heif"
          >
            Upload File(s)
          </button>

          <button
            type="button"
            className="primary-button"
            onClick={
              handleGetResponse
            }
            disabled={isLoading}
          >
            {isLoading
              ? "Processing..."
              : "Get Response"}
          </button>

        </div>

        {/* Editable Response */}

        {response && (
          <div className="response-container">

            <div className="response-title">
              Response
            </div>

            <textarea
              className="response-editor"
              value={response}
              onChange={(event) =>
                setResponse(
                  event.target.value
                )
              }
              spellCheck={false}
            />

            <div className="download-container">

              <button
                type="button"
                className="primary-button"
                onClick={
                  handleDownloadDocx
                }
              >
                Download
              </button>

            </div>

          </div>
        )}

      </main>

    </div>
  );
}

export default App;