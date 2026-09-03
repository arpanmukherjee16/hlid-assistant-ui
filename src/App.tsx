import { useEffect, useRef, useState } from "react";
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

interface ConversationFile {
  name: string;
  contentType: string;
  size: number;
}

interface ConversationMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  files?: ConversationFile[];
}

interface StoredConversationMessage
  extends ConversationMessage {
  processedFiles?: ProcessedFile[];
}

interface LogicAppPayload {
  prompt: string;
  files: ProcessedFile[];
  conversation: ConversationMessage[];
}

function App() {
  const fileInputRef =
    useRef<HTMLInputElement>(null);

  const chatEndRef =
    useRef<HTMLDivElement>(null);

  const [prompt, setPrompt] =
    useState("");

  const [selectedFiles, setSelectedFiles] =
    useState<File[]>([]);

  const [processedFiles, setProcessedFiles] =
    useState<ProcessedFile[]>([]);

  const [messages, setMessages] =
    useState<StoredConversationMessage[]>(
      []
    );

  const [isLoading, setIsLoading] =
    useState(false);

  const [editingResponseId, setEditingResponseId] =
    useState<string | null>(null);

  const [editingResponseText, setEditingResponseText] =
    useState("");

  const [showInstructions, setShowInstructions] =
    useState(false);

  /*
   * Automatically scroll to the latest
   * message whenever the conversation changes.
   */
  useEffect(() => {
    if (messages.length > 0) {
      chatEndRef.current?.scrollIntoView({
        behavior: "smooth",
      });
    }
  }, [messages, isLoading]);

  /*
   * Generate a unique message ID.
   */
  const createMessageId = () => {
    return `${Date.now()}-${Math.random()
      .toString(36)
      .substring(2, 10)}`;
  };

  /*
   * Convert a browser File into Base64.
   */
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
                `Unable to extract Base64 content: ${file.name}`
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

  /*
   * Open the browser file picker.
   */
  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  /*
   * Handle prompt changes.
   */
  const handlePromptChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>
  ) => {
    setPrompt(event.target.value);
  };

  /*
   * Handle uploaded files.
   *
   * The important part is that the Base64
   * representation is retained in processedFiles.
   */
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
            file.type ||
            "application/octet-stream",
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

      alert(
        "Unable to process one or more selected files."
      );
    }

    /*
     * Allows the same file to be selected
     * again later.
     */
    event.target.value = "";
  };

  /*
   * Remove one file from the current composer.
   */
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

  /*
   * Clear files currently attached to
   * the composer.
   */
  const clearComposerFiles = () => {
    setSelectedFiles([]);
    setProcessedFiles([]);

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  /*
   * Convert full ProcessedFile objects
   * into lightweight metadata for the
   * visible conversation.
   */
  const getConversationFiles = (
    files: ProcessedFile[]
  ): ConversationFile[] => {
    return files.map(
      (file) => ({
        name: file.name,
        contentType:
          file.contentType,
        size: file.size,
      })
    );
  };

  /*
   * Convert the internal message format
   * into the public conversation format
   * expected by the Logic App.
   *
   * Base64 file contents are deliberately
   * NOT sent inside conversation history.
   */
  const getPublicConversation = (
    conversationMessages: StoredConversationMessage[]
  ): ConversationMessage[] => {
    return conversationMessages.map(
      (message) => ({
        id: message.id,
        role: message.role,
        content: message.content,
        files: message.files,
      })
    );
  };

  /*
   * Call the Logic App and poll its
   * asynchronous response.
   */
  const callLogicApp = async (
    payload: LogicAppPayload
  ): Promise<string> => {
    const logicAppUrl =
      import.meta.env.VITE_LOGIC_APP_URL;

    if (!logicAppUrl) {
      throw new Error(
        "Logic App URL is not configured."
      );
    }

    const result = await fetch(
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

    const location =
      result.headers.get(
        "Location"
      );

    if (result.status !== 202) {
      const errorText =
        await result.text();

      throw new Error(
        `Logic App returned HTTP ${result.status}: ${errorText}`
      );
    }

    if (!location) {
      throw new Error(
        "Logic App returned 202 but no Location header was provided."
      );
    }

    while (true) {
      await new Promise(
        (resolve) =>
          setTimeout(
            resolve,
            3000
          )
      );

      const pollResult =
        await fetch(
          location,
          {
            method: "GET",
          }
        );

      const pollText =
        await pollResult.text();

      if (
        pollResult.status ===
        202
      ) {
        continue;
      }

      if (
        pollResult.status ===
        200
      ) {
        return (
          pollText ||
          "No response was returned by the Logic App."
        );
      }

      throw new Error(
        `Logic App polling failed with HTTP ${pollResult.status}: ${pollText}`
      );
    }
  };

  /*
   * Determine which files belong to a
   * particular previous user message.
   *
   * The actual Base64 content is stored
   * in processedFiles on that message.
   */
  const getFilesForMessage = (
    message:
      StoredConversationMessage
  ): ProcessedFile[] => {
    return (
      message.processedFiles ??
      []
    );
  };

  /*
   * Find the most recent uploaded files
   * in the conversation.
   *
   * This is used for a follow-up request
   * such as:
   *
   * "Now identify the gaps."
   *
   * after:
   *
   * "Read this document."
   */
  const getRelevantFilesForFollowUp = (
    conversation:
      StoredConversationMessage[]
  ): ProcessedFile[] => {
    /*
     * Search backwards for the latest
     * user message that contained files.
     */
    for (
      let index =
        conversation.length - 1;
      index >= 0;
      index--
    ) {
      const message =
        conversation[index];

      if (
        message.role ===
          "user" &&
        message.processedFiles &&
        message.processedFiles.length >
          0
      ) {
        return message.processedFiles;
      }
    }

    return [];
  };

  /*
   * Submit a completely new request.
   */
  const handleGetResponse =
    async () => {
      if (!prompt.trim()) {
        alert(
          "Please enter an instruction or query."
        );
        return;
      }

      if (isLoading) {
        return;
      }

      const userPrompt =
        prompt.trim();

      /*
       * Files currently attached to the
       * composer belong to this new
       * user message.
       */
      const filesForThisRequest =
        [...processedFiles];

      const userMessage:
        StoredConversationMessage = {
          id: createMessageId(),

          role: "user",

          content:
            userPrompt,

          files:
            filesForThisRequest.length >
            0
              ? getConversationFiles(
                  filesForThisRequest
                )
              : [],

          /*
           * IMPORTANT:
           * Keep the actual Base64 files
           * attached to this message.
           */
          processedFiles:
            filesForThisRequest,
        };

      /*
       * For a follow-up request where the
       * user didn't upload a new file,
       * automatically reuse the most recent
       * relevant uploaded files.
       */
      let filesForRequest =
        filesForThisRequest;

      if (
        filesForRequest.length ===
        0
      ) {
        filesForRequest =
          getRelevantFilesForFollowUp(
            messages
          );
      }

      const conversationWithUserMessage =
        [
          ...messages,
          userMessage,
        ];

      setMessages(
        conversationWithUserMessage
      );

      setPrompt("");

      clearComposerFiles();

      setIsLoading(true);

      try {
        const payload:
          LogicAppPayload = {
          prompt:
            userPrompt,

          files:
            filesForRequest,

          conversation:
            getPublicConversation(
              conversationWithUserMessage
            ),
        };

        const finalResponse =
          await callLogicApp(
            payload
          );

        const assistantMessage:
          StoredConversationMessage = {
          id: createMessageId(),

          role: "assistant",

          content:
            finalResponse ||
            "No response was returned by the Logic App.",
        };

        setMessages(
          (previousMessages) => [
            ...previousMessages,
            assistantMessage,
          ]
        );
      } catch (error) {
        console.error(
          "Logic App request failed:",
          error
        );

        const errorMessage:
          StoredConversationMessage = {
          id: createMessageId(),

          role: "assistant",

          content:
            error instanceof Error
              ? error.message
              : "An unexpected error occurred.",
        };

        setMessages(
          (previousMessages) => [
            ...previousMessages,
            errorMessage,
          ]
        );
      } finally {
        setIsLoading(false);
      }
    };

  /*
   * Regenerate an existing assistant
   * response without creating another
   * user message.
   */
  const handleRegenerate = async (
    assistantMessageId: string
  ) => {
    if (isLoading) {
      return;
    }

    const assistantIndex =
      messages.findIndex(
        (message) =>
          message.id ===
            assistantMessageId &&
          message.role ===
            "assistant"
      );

    if (assistantIndex <= 0) {
      return;
    }

    const userMessage =
      messages[
        assistantIndex - 1
      ];

    if (
      !userMessage ||
      userMessage.role !==
        "user"
    ) {
      return;
    }

    setIsLoading(true);

    try {
      /*
       * Reuse the exact files that were
       * attached to the original request.
       */
      const filesForRequest =
        getFilesForMessage(
          userMessage
        );

      /*
       * Everything before the assistant
       * response being regenerated becomes
       * the conversation context.
       */
      const conversationBeforeAssistant =
        messages.slice(
          0,
          assistantIndex
        );

      const payload:
        LogicAppPayload = {
        prompt:
          userMessage.content,

        files:
          filesForRequest,

        conversation:
          getPublicConversation(
            conversationBeforeAssistant
          ),
      };

      const finalResponse =
        await callLogicApp(
          payload
        );

      setMessages(
        (previousMessages) =>
          previousMessages.map(
            (message) =>
              message.id ===
              assistantMessageId
                ? {
                    ...message,

                    content:
                      finalResponse ||
                      "No response was returned by the Logic App.",
                  }
                : message
          )
      );
    } catch (error) {
      console.error(
        "Regeneration failed:",
        error
      );

      setMessages(
        (previousMessages) =>
          previousMessages.map(
            (message) =>
              message.id ===
              assistantMessageId
                ? {
                    ...message,

                    content:
                      error instanceof
                      Error
                        ? error.message
                        : "Unable to regenerate the response.",
                  }
                : message
          )
      );
    } finally {
      setIsLoading(false);
    }
  };

  /*
   * Edit the generated ASSISTANT response
   * before downloading it.
   */
  const handleEditResponse = (
    message: StoredConversationMessage
  ) => {
    setEditingResponseId(message.id);
    setEditingResponseText(message.content);
  };

  /*
   * Save the edited assistant response.
   */
  const handleSaveResponseEdit = (
    messageId: string
  ) => {
    if (!editingResponseText.trim()) {
      alert("The response cannot be empty.");
      return;
    }

    setMessages(
      (previousMessages) =>
        previousMessages.map(
          (message) =>
            message.id === messageId
              ? {
                  ...message,
                  content:
                    editingResponseText,
                }
              : message
        )
    );

    setEditingResponseId(null);
    setEditingResponseText("");
  };

  /*
   * Cancel response editing without
   * changing the generated response.
   */
  const handleCancelResponseEdit = () => {
    setEditingResponseId(null);
    setEditingResponseText("");
  };

  /*
   * Download an individual assistant
   * response as DOCX.
   */
  const handleDownloadDocx =
    async (
      message:
        StoredConversationMessage
    ) => {
      if (
        !message.content.trim()
      ) {
        alert(
          "There is no response to download."
        );
        return;
      }

      try {
        const lines =
          message.content.split(
            /\r?\n/
          );

        const paragraphs =
          lines.map(
            (line) =>
              new Paragraph({
                alignment:
                  AlignmentType.JUSTIFIED,

                children: [
                  new TextRun({
                    text: line,
                    size: 22,
                  }),
                ],
              })
          );

        const docxDocument =
          new Document({
            sections: [
              {
                properties: {},
                children:
                  paragraphs,
              },
            ],
          });

        const blob =
          await Packer.toBlob(
            docxDocument
          );

        const downloadUrl =
          URL.createObjectURL(
            blob
          );

        const link =
          window.document.createElement(
            "a"
          );

        link.href =
          downloadUrl;

        link.download =
          "HLID_Response.docx";

        window.document.body.appendChild(
          link
        );

        link.click();

        window.document.body.removeChild(
          link
        );

        URL.revokeObjectURL(
          downloadUrl
        );
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

  /*
   * Enter = send
   * Shift + Enter = new line
   */
  const handlePromptKeyDown = (
    event:
      React.KeyboardEvent<HTMLTextAreaElement>
  ) => {
    if (
      event.key === "Enter" &&
      !event.shiftKey
    ) {
      event.preventDefault();

      void handleGetResponse();
    }
  };

  /*
   * Start a completely new chat.
   */
  const handleStartNewChat = () => {
    setMessages([]);
    setPrompt("");
    setProcessedFiles([]);
    setSelectedFiles([]);
    setEditingResponseId(null);
    setEditingResponseText("");
    setShowInstructions(false);

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }

    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  };

  return (
    <div className="app-container">

      {/* =================================================
          HEADER
          ================================================= */}

      <header className="app-header">

        <div className="app-branding">
          <h1>
            Welcome at HLID Assistant
          </h1>

         </div>

        <div className="header-actions">
          <button
            type="button"
            className="info-button"
            onClick={() =>
              setShowInstructions(true)
            }
            title="How HLID Assistant generates responses"
            aria-label="How HLID Assistant generates responses"
          >
            i
          </button>

          <button
            type="button"
            className="new-chat-button"
            onClick={handleStartNewChat}
            disabled={isLoading}
            title="Start a new chat"
          >
            + Start New Chat
          </button>
        </div>

      </header>


      {/* =================================================
          CHAT
          ================================================= */}

      <main className="chat-page">

        <div className="conversation-container">

          {messages.length ===
            0 && (
            <div className="welcome-panel">

              <h2>
                How can I help you?
              </h2>

              <p>
                Ask a question, provide
                an instruction, or upload
                a document for analysis.
              </p>

            </div>
          )}


          {messages.map(
            (
              message,
              index
            ) => (

              <div
                key={
                  message.id
                }
                className={`message-row ${
                  message.role ===
                  "user"
                    ? "user-message-row"
                    : "assistant-message-row"
                }`}
              >

                <div
                  className={`message-bubble ${
                    message.role ===
                    "user"
                      ? "user-message-bubble"
                      : "assistant-message-bubble"
                  }`}
                >

                  <div className="message-role">
                    {message.role ===
                    "user"
                      ? "You"
                      : "HLID Assistant"}
                  </div>


                  {/* ======================================
                      FILE ATTACHMENTS
                      ====================================== */}

                  {message.files &&
                    message.files
                      .length >
                      0 && (

                      <div className="message-files">

                        {message.files.map(
                          (
                            file,
                            fileIndex
                          ) => (

                            <div
                              className="message-file"
                              key={`${file.name}-${fileIndex}`}
                            >

                              <span className="message-file-icon">
                                📎
                              </span>

                              <span className="message-file-name">
                                {
                                  file.name
                                }
                              </span>

                            </div>

                          )
                        )}

                      </div>

                  )}


                  {/* ======================================
                      MESSAGE CONTENT
                      ====================================== */}

                  {message.role ===
                    "assistant" &&
                    editingResponseId ===
                      message.id ? (

                    <textarea
                      className="response-edit-box"
                      value={
                        editingResponseText
                      }
                      onChange={(
                        event
                      ) =>
                        setEditingResponseText(
                          event.target.value
                        )
                      }
                      aria-label="Edit generated response"
                    />

                  ) : (

                    <div className="message-content">
                      {
                        message.content
                      }
                    </div>

                  )}


                  {/* ======================================
                      ASSISTANT ACTIONS
                      ====================================== */}

                  {message.role ===
                    "assistant" && (

                    <div className="assistant-actions">

                      {editingResponseId ===
                        message.id ? (

                        <>
                          <button
                            type="button"
                            className="message-action-button response-save-button"
                            onClick={() =>
                              handleSaveResponseEdit(
                                message.id
                              )
                            }
                            disabled={
                              isLoading
                            }
                            title="Save edited response"
                          >
                            ✓ Save Changes
                          </button>

                          <button
                            type="button"
                            className="message-action-button"
                            onClick={
                              handleCancelResponseEdit
                            }
                            disabled={
                              isLoading
                            }
                            title="Cancel response editing"
                          >
                            ✕ Cancel
                          </button>
                        </>

                      ) : (

                        <>
                          <button
                            type="button"
                            className="message-action-button"
                            onClick={() =>
                              void handleRegenerate(
                                message.id
                              )
                            }
                            disabled={
                              isLoading
                            }
                            title="Regenerate response"
                          >
                            ↻ Regenerate
                          </button>

                          <button
                            type="button"
                            className="message-action-button"
                            onClick={() =>
                              handleEditResponse(
                                message
                              )
                            }
                            disabled={
                              isLoading
                            }
                            title="Edit generated response before downloading"
                          >
                            ✎ Edit Response
                          </button>

                          <button
                            type="button"
                            className="message-action-button"
                            onClick={() =>
                              void handleDownloadDocx(
                                message
                              )
                            }
                            disabled={
                              isLoading
                            }
                            title="Download response as Word document"
                          >
                            ↓ Download
                          </button>
                        </>

                      )}

                    </div>

                  )}

                </div>

              </div>

            )
          )}


          {/* ==============================================
              LOADING MESSAGE
              ============================================== */}

          {isLoading && (

            <div className="message-row assistant-message-row">

              <div className="message-bubble assistant-message-bubble">

                <div className="message-role">
                  HLID Assistant
                </div>

                <div className="typing-indicator">

                  <span></span>
                  <span></span>
                  <span></span>

                </div>

              </div>

            </div>

          )}


          <div
            ref={chatEndRef}
            className="chat-end"
          />

        </div>

      </main>


      {/* =================================================
          INSTRUCTIONS MODAL
          ================================================= */}

      {showInstructions && (
        <div
          className="instructions-overlay"
          onClick={() =>
            setShowInstructions(false)
          }
        >
          <div
            className="instructions-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="instructions-title"
            onClick={(event) =>
              event.stopPropagation()
            }
          >
            <div className="instructions-header">
              <div>
                <h2 id="instructions-title">
                  How HLID Assistant Generates Responses
                </h2>
                <p>
                  The response style is determined primarily by your request,
                  while the supplied document content is used as the factual source.
                </p>
              </div>

              <button
                type="button"
                className="instructions-close-button"
                onClick={() =>
                  setShowInstructions(false)
                }
                aria-label="Close instructions"
                title="Close"
              >
                ×
              </button>
            </div>

            <div className="instructions-content">

              <section>
                <h3>1. Direct question or normal request</h3>
                <p>
                  If you ask a direct question, the assistant answers that
                  question directly and proportionately. It does not automatically
                  turn the answer into a full HLID section.
                </p>
              </section>

              <section>
                <h3>2. Pasted document-like content</h3>
                <p>
                  If you paste meeting minutes, an email, requirements, or other
                  business/technical source content into the prompt, the supplied
                  content is treated as the source. When you ask to generate or
                  summarize content, the assistant rewrites the substantive
                  information as professional HLID-ready content.
                </p>
              </section>

              <section>
                <h3>3. Uploaded documents or images</h3>
                <p>
                  Uploaded documents and images are analyzed for their business
                  and technical content. The information obtained from the current
                  request is treated as the authoritative source for facts from
                  those files.
                </p>
              </section>

              <section>
                <h3>4. Generate, rewrite, or summarize</h3>
                <p>
                  When you ask to generate, rewrite, prepare, or summarize
                  content, the assistant writes the actual business or technical
                  content rather than describing what the source contains.
                  The result is intended to be directly reusable in HLID.
                  When enough substantive information is available, generated
                  content normally contains at least three meaningful paragraphs
                  or clearly separated content blocks.
                </p>
              </section>

              <section>
                <h3>5. Meeting minutes and action items</h3>
                <p>
                  Meeting minutes are treated as source material, not as a
                  template for the final answer. Decisions, requirements,
                  technical behaviour, workflows, and agreed implementation
                  details are retained. Action items, owners, responsibilities,
                  pending items, and next steps are not automatically included
                  unless you explicitly ask for them.
                </p>
              </section>

              <section>
                <h3>6. Terminology and factual accuracy</h3>
                <p>
                  Important names, numbers, system names, message names,
                  interfaces, workflows, business rules, and technical details
                  are preserved wherever possible. The assistant does not invent
                  or silently fill missing information with unsupported facts.
                </p>
              </section>

              <section>
                <h3>7. Multiple uploaded files</h3>
                <p>
                  When multiple files are supplied, their relevant analyzed
                  information is considered together and consolidated into one
                  response unless you explicitly ask for a document-by-document
                  comparison.
                </p>
              </section>

              <section>
                <h3>8. HLID writing style</h3>
                <p>
                  Generated documentation uses professional paragraphs as the
                  primary format. Meaningful headings or sub-headings are used
                  when they improve readability. Generic headings, unnecessary
                  bullet lists, meeting-report formats, and meta-document wording
                  are avoided.
                </p>
              </section>

            </div>
          </div>
        </div>
      )}

      {/* =================================================
          COMPOSER
          ================================================= */}

      <footer className="composer-area">


        {/* ================================================
            SELECTED FILES
            ================================================ */}

        {selectedFiles.length >
          0 && (

          <div className="composer-files">

            {selectedFiles.map(
              (
                file,
                index
              ) => (

                <div
                  className="composer-file"
                  key={`${file.name}-${index}`}
                >

                  <span className="composer-file-name">
                    📎{" "}
                    {
                      file.name
                    }
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


        {/* ================================================
            COMPOSER
            ================================================ */}

        <div className="composer">

          <textarea
            className="query-box"
            placeholder="Message HLID Assistant..."
            value={prompt}
            onChange={
              handlePromptChange
            }
            onKeyDown={
              handlePromptKeyDown
            }
            disabled={
              isLoading
            }
          />


          <div className="composer-controls">


            {/* ============================================
                FILE INPUT
                ============================================ */}

            <input
              ref={
                fileInputRef
              }
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
                .ppt,
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
                display:
                  "none",
              }}
              onChange={
                handleFileChange
              }
            />


            {/* ============================================
                UPLOAD
                ============================================ */}

            <div className="upload-button-wrapper">
              <button
                type="button"
                className="composer-upload-button"
                onClick={
                  handleUploadClick
                }
                disabled={
                  isLoading
                }
              >
                + Upload File(s)
              </button>

              <div className="upload-file-types-tooltip">
                <strong>Allowed file types:</strong>
                <span>
                  .c, .cpp, .cs, .css, .doc, .docx, .go, .html, .java,
                  .js, .json, .md, .pdf, .php, .pptx, .py, .rb, .sh,
                  .tex, .ts, .txt, .jpg, .jpeg, .png, .bmp, .tiff, .heif
                </span>
              </div>
            </div>


            {/* ============================================
                SEND
                ============================================ */}

            <button
              type="button"
              className="composer-send-button"
              onClick={() =>
                void handleGetResponse()
              }
              disabled={
                isLoading ||
                !prompt.trim()
              }
              title="Send"
            >
              {isLoading
                ? "Processing..."
                : "Send"}
            </button>

          </div>

        </div>


        {/* ================================================
            HINT
            ================================================ */}

        <div className="composer-hint">
          Enter to send · Shift +
          Enter for a new line
        </div>

      </footer>

    </div>
  );
}

export default App;
