import { useRef, useState } from "react";
import "./App.css";

const API_URL = "/api";

function App() {
  const [selfFile, setSelfFile] = useState<File | null>(null);
  const [superiorFile, setSuperiorFile] =
    useState<File | null>(null);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");

  const selfInputRef =
    useRef<HTMLInputElement>(null);

  const superiorInputRef =
    useRef<HTMLInputElement>(null);

  const canGenerate =
    selfFile !== null &&
    superiorFile !== null &&
    !loading;

  const handleSelfFile = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file =
      event.target.files?.[0] ?? null;

    setSelfFile(file);
    setError("");
    setSuccess("");
  };

  const handleSuperiorFile = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file =
      event.target.files?.[0] ?? null;

    setSuperiorFile(file);
    setError("");
    setSuccess("");
  };

  const handleGenerate = async () => {
    if (!selfFile || !superiorFile) {
      return;
    }

    setLoading(true);
    setError("");
    setSuccess("");

    try {
      const formData = new FormData();

      formData.append(
        "self_assessment",
        selfFile
      );

      formData.append(
        "superior_assessment",
        superiorFile
      );

      const response = await fetch(
        `${API_URL}/generate`,
        {
          method: "POST",
          body: formData,
        }
      );

      if (!response.ok) {
        let message =
          "Gagal memproses assessment.";

        try {
          const errorData =
            await response.json();

          if (
            typeof errorData.detail ===
            "string"
          ) {
            message = errorData.detail;
          } else if (
            errorData.detail?.message
          ) {
            message =
              errorData.detail.message;
          } else if (
            errorData.detail?.error
          ) {
            message =
              errorData.detail.error;
          }
        } catch {
          // Ignore JSON parsing error.
        }

        throw new Error(message);
      }

      const blob = await response.blob();

      if (blob.size === 0) {
        throw new Error(
          "File hasil dari server kosong."
        );
      }

      const downloadUrl =
        window.URL.createObjectURL(blob);

      const link =
        document.createElement("a");

      link.href = downloadUrl;

      link.download =
        "Learning_Agility_Reports.zip";

      document.body.appendChild(link);
      link.click();
      link.remove();

      window.URL.revokeObjectURL(
        downloadUrl
      );

      setSuccess(
        "Report berhasil dibuat dan di-download."
      );
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Terjadi error saat memproses assessment."
      );
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setSelfFile(null);
    setSuperiorFile(null);
    setError("");
    setSuccess("");

    if (selfInputRef.current) {
      selfInputRef.current.value = "";
    }

    if (superiorInputRef.current) {
      superiorInputRef.current.value = "";
    }
  };

  return (
    <main className="app">

      {/* =====================================================
          MAIN CARD
      ===================================================== */}

      <section className="card">

        {/* HEADER */}

        <header className="header">

          <div className="brand-icon">
            <span></span>
            <span></span>
          </div>

          <div>
            <h1>
              Learning <strong>Agility</strong>
            </h1>

            <p>
              Learning Agility Assessment System
            </p>
          </div>

        </header>

        {/* DESCRIPTION */}

        <div className="description">
          Upload hasil assessment untuk
          menghasilkan Learning Agility Report
          secara otomatis.
        </div>

        {/* ===================================================
            SELF ASSESSMENT
        =================================================== */}

        <div className="form-group">

          <label>
            <span className="label-icon">
              ♙
            </span>

            Self Assessment
          </label>

          <div
            className={`upload-area ${
              selfFile ? "has-file" : ""
            }`}
            onClick={() => {
              if (!loading) {
                selfInputRef.current?.click();
              }
            }}
          >

            <input
              ref={selfInputRef}
              type="file"
              accept=".xlsx,.xls"
              disabled={loading}
              onChange={handleSelfFile}
            />

            {!selfFile ? (
              <div className="upload-content">

                <div className="upload-icon">
                  ↑
                </div>

                <div className="upload-text">
                  <strong>
                    Choose File
                  </strong>

                  <small>
                    Excel file (.xlsx, .xls)
                  </small>
                </div>

                <div className="document-icon">
                  ▱
                </div>

              </div>
            ) : (
              <div className="selected-file">

                <div className="excel-icon">
                  XLS
                </div>

                <div className="file-info">

                  <strong>
                    {selfFile.name}
                  </strong>

                  <small>
                    Self Assessment
                  </small>

                </div>

                <div className="file-check">
                  ✓
                </div>

              </div>
            )}

          </div>

        </div>

        {/* ===================================================
            SUPERIOR ASSESSMENT
        =================================================== */}

        <div className="form-group">

          <label>
            <span className="label-icon">
              ♙
            </span>

            Superior Assessment
          </label>

          <div
            className={`upload-area ${
              superiorFile ? "has-file" : ""
            }`}
            onClick={() => {
              if (!loading) {
                superiorInputRef.current?.click();
              }
            }}
          >

            <input
              ref={superiorInputRef}
              type="file"
              accept=".xlsx,.xls"
              disabled={loading}
              onChange={
                handleSuperiorFile
              }
            />

            {!superiorFile ? (
              <div className="upload-content">

                <div className="upload-icon">
                  ↑
                </div>

                <div className="upload-text">
                  <strong>
                    Choose File
                  </strong>

                  <small>
                    Excel file (.xlsx, .xls)
                  </small>
                </div>

                <div className="document-icon">
                  ▱
                </div>

              </div>
            ) : (
              <div className="selected-file">

                <div className="excel-icon">
                  XLS
                </div>

                <div className="file-info">

                  <strong>
                    {superiorFile.name}
                  </strong>

                  <small>
                    Superior Assessment
                  </small>

                </div>

                <div className="file-check">
                  ✓
                </div>

              </div>
            )}

          </div>

        </div>

        {/* ===================================================
            STATUS
        =================================================== */}

        {loading && (
          <div className="status loading">
            <span className="status-spinner"></span>

            <div>
              <strong>
                Processing assessment...
              </strong>

              <small>
                Please wait while the report
                is being generated.
              </small>
            </div>
          </div>
        )}

        {success && !loading && (
          <div className="status success">
            <span className="status-check">
              ✓
            </span>

            <div>
              <strong>
                Report berhasil dibuat
              </strong>

              <small>
                File ZIP berhasil di-download.
              </small>
            </div>
          </div>
        )}

        {error && !loading && (
          <div className="status error">
            <span className="status-check">
              !
            </span>

            <div>
              <strong>
                Processing gagal
              </strong>

              <small>
                {error}
              </small>
            </div>
          </div>
        )}

        {/* ===================================================
            GENERATE BUTTON
        =================================================== */}

        <button
          type="button"
          className="generate-button"
          disabled={!canGenerate}
          onClick={handleGenerate}
        >
          {loading ? (
            <>
              <span className="button-spinner"></span>
              Processing...
            </>
          ) : (
            <>
              <span className="sparkle">
                ✦
              </span>

              Generate Report

              <span className="arrow">
                →
              </span>
            </>
          )}
        </button>

        {/* RESET */}

        {(selfFile ||
          superiorFile ||
          success ||
          error) &&
          !loading && (
            <button
              type="button"
              className="reset-button"
              onClick={handleReset}
            >
              Reset
            </button>
          )}

        {/* FOOTER */}

        <div className="footer">
          <span>◇</span>

          Pastikan kedua file berformat
          Excel (.xlsx atau .xls)
        </div>

      </section>
    </main>
  );
}

export default App;

