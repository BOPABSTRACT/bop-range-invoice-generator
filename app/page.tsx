'use client'

import { useState, useRef } from 'react'

const LOGO = "https://i.imgur.com/szjzoxt.png"

// Vercel serverless functions have a 4.5 MB request body limit.
// We use 4.4 MB to leave a little headroom for form-data overhead.
const MAX_UPLOAD_BYTES = 4_400_000

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

export default function Home() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [passwordError, setPasswordError] = useState(false)
  const [excelFiles, setExcelFiles] = useState<File[]>([])
  const [receiptFiles, setReceiptFiles] = useState<File[]>([])
  const [emailFile, setEmailFile] = useState<File | null>(null)
  const [invoiceDate, setInvoiceDate] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'done' | 'error'>('idle')
  const [message, setMessage] = useState('')
  const excelRef = useRef<HTMLInputElement>(null)
  const receiptRef = useRef<HTMLInputElement>(null)
  const emailRef = useRef<HTMLInputElement>(null)

  // Live-computed total upload size and over-limit flag.
  const totalBytes =
    excelFiles.reduce((sum, f) => sum + f.size, 0) +
    receiptFiles.reduce((sum, f) => sum + f.size, 0) +
    (emailFile?.size || 0)
  const overLimit = totalBytes > MAX_UPLOAD_BYTES
  const hasAnyFile = excelFiles.length > 0 || receiptFiles.length > 0 || emailFile !== null

  const handlePasswordSubmit = () => {
    if (passwordInput === 'BOP2026') {
      setAuthenticated(true)
      setPasswordError(false)
    } else {
      setPasswordError(true)
    }
  }

  const handleGenerate = async () => {
    if (excelFiles.length === 0) {
      setMessage('Please upload at least one Excel file.')
      setStatus('error')
      return
    }
    if (!emailFile) {
      setMessage('Please upload the billing email PDF.')
      setStatus('error')
      return
    }
    if (!invoiceDate.trim()) {
      setMessage('Please enter an invoice date.')
      setStatus('error')
      return
    }
    if (overLimit) {
      setMessage(
        `Your files total ${formatBytes(totalBytes)} — too large to upload. ` +
        `Combined size must be under ${formatBytes(MAX_UPLOAD_BYTES)}. ` +
        `The receipts PDF is usually the heaviest — try compressing it at iLovePDF.com or SmallPDF.com.`
      )
      setStatus('error')
      return
    }
    setStatus('loading')
    setMessage('Generating invoices...')

    const formData = new FormData()
    excelFiles.forEach(f => formData.append('excel', f))
    receiptFiles.forEach(f => formData.append('receipts', f))
    formData.append('email', emailFile)
    formData.append('invoiceDate', invoiceDate.trim())

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: formData })
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.error || 'Generation failed')
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `range-invoices-${Date.now()}.zip`
      a.click()
      URL.revokeObjectURL(url)
      setStatus('done')
      setMessage(`✅ ${excelFiles.length} invoice${excelFiles.length !== 1 ? 's' : ''} generated and downloaded!`)
    } catch (err: unknown) {
      setStatus('error')
      setMessage(err instanceof Error ? err.message : 'Something went wrong.')
    }
  }

  if (!authenticated) {
    return (
      <main style={{
        minHeight: '100vh', background: '#0f1117', fontFamily: "'Georgia', serif",
        color: '#e8e0d0', display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}>
        <div style={{
          background: '#0d0f14', border: '1px solid #2a2a3a', borderRadius: 12,
          padding: '48px 40px', width: '100%', maxWidth: 400, textAlign: 'center',
        }}>
          <img src={LOGO} alt="BOP Logo"
            style={{ width: 140, height: 140, objectFit: 'contain', margin: '0 auto 24px', display: 'block' }} />
          <div style={{ fontSize: 20, fontWeight: 600, color: '#c8a96e', marginBottom: 4 }}>
            Range Resources Title
          </div>
          <div style={{ fontSize: 12, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase', marginBottom: 32 }}>
            Invoice Generator
          </div>
          <input
            type="password" placeholder="Enter password" value={passwordInput}
            onChange={e => { setPasswordInput(e.target.value); setPasswordError(false) }}
            onKeyDown={e => e.key === 'Enter' && handlePasswordSubmit()}
            style={{
              width: '100%', padding: '12px 16px', background: '#0f1117',
              border: `1px solid ${passwordError ? '#8b2020' : '#2a2a3a'}`,
              borderRadius: 6, color: '#e8e0d0', fontSize: 15,
              fontFamily: "'Georgia', serif", boxSizing: 'border-box', marginBottom: 12, outline: 'none',
            }}
          />
          {passwordError && (
            <div style={{ color: '#e07070', fontSize: 13, marginBottom: 12 }}>
              Incorrect password. Please try again.
            </div>
          )}
          <button onClick={handlePasswordSubmit} style={{
            width: '100%', padding: '12px 32px',
            background: 'linear-gradient(135deg, #c8a96e, #8b6914)',
            color: '#fff', border: 'none', borderRadius: 6, fontSize: 15,
            fontFamily: "'Georgia', serif", cursor: 'pointer', letterSpacing: '0.04em',
          }}>Enter</button>
        </div>
      </main>
    )
  }

  return (
    <main style={{ minHeight: '100vh', background: '#0f1117', fontFamily: "'Georgia', serif", color: '#e8e0d0' }}>
      <header style={{
        borderBottom: '1px solid #2a2a3a', padding: '16px 48px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#0d0f14',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <img src={LOGO} alt="BOP Logo" style={{ width: 52, height: 52, objectFit: 'contain', flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '0.04em', color: '#c8a96e' }}>
              Range Resources Title
            </div>
            <div style={{ fontSize: 11, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase' }}>
              Invoice Generator
            </div>
          </div>
        </div>
        <a href="/help" style={{
          color: '#c8a96e', fontSize: 13, textDecoration: 'none',
          border: '1px solid #333', padding: '6px 14px', borderRadius: 4, letterSpacing: '0.04em',
        }}>User Guide</a>
      </header>

      <div style={{ maxWidth: 820, margin: '0 auto', padding: '48px 32px' }}>
        <div style={{ marginBottom: 48 }}>
          <h1 style={{ fontSize: 36, fontWeight: 400, color: '#e8e0d0', margin: '0 0 12px 0', letterSpacing: '-0.01em', lineHeight: 1.2 }}>
            Generate Invoices
          </h1>
          <p style={{ color: '#888', fontSize: 15, margin: 0, lineHeight: 1.6 }}>
            Upload your Excel invoice files, any receipt PDFs, and the billing email PDF.
            One complete PDF invoice will be generated per Excel file.
          </p>
        </div>

        <Section number="1" title="Upload Excel Invoice Files">
          <MultiUploadBox
            label="Drop one or more .xlsx invoice files here or click to browse"
            accept=".xlsx,.xls"
            files={excelFiles}
            onChange={e => setExcelFiles(Array.from(e.target.files || []))}
            inputRef={excelRef}
            icon="📊"
          />
        </Section>

        <Section number="2" title="Upload Receipt PDFs (Optional)">
          <MultiUploadBox
            label="Drop receipt PDFs here or click to browse — optional"
            accept=".pdf"
            files={receiptFiles}
            onChange={e => setReceiptFiles(Array.from(e.target.files || []))}
            inputRef={receiptRef}
            icon="🧾"
          />
          <div style={{ fontSize: 12, color: '#666', marginTop: 8 }}>
            Receipts are matched to invoices by invoice number. Each receipt will be appended to its matching invoice.
          </div>
        </Section>

        <Section number="3" title="Upload Billing Email PDF">
          <SingleUploadBox
            label="Drop the billing email PDF here or click to browse"
            accept=".pdf"
            file={emailFile}
            onChange={e => setEmailFile(e.target.files?.[0] || null)}
            inputRef={emailRef}
            icon="📧"
          />
        </Section>

        <Section number="4" title="Invoice Date">
          <div style={{ fontSize: 13, color: '#888', marginBottom: 10 }}>
            Enter the date to appear on all invoices (e.g. May 12, 2026 or 5/12/2026)
          </div>
          <input
            type="text"
            placeholder="e.g. May 12, 2026"
            value={invoiceDate}
            onChange={e => setInvoiceDate(e.target.value)}
            style={{
              width: '100%', padding: '12px 16px', background: '#0d0f14',
              border: `1px solid ${!invoiceDate && status === 'error' ? '#8b2020' : '#2a2a3a'}`,
              borderRadius: 6, color: '#e8e0d0', fontSize: 15,
              fontFamily: "'Georgia', serif", boxSizing: 'border-box', outline: 'none',
            }}
          />
        </Section>

        <Section number="5" title="Generate Invoices">
          {hasAnyFile && (
            <div style={{
              marginBottom: 16,
              padding: '12px 16px',
              borderRadius: 6,
              background: overLimit ? 'rgba(200,60,60,0.08)' : 'rgba(200,169,110,0.04)',
              border: `1px solid ${overLimit ? '#8b2020' : '#2a2a3a'}`,
              fontSize: 13,
              color: overLimit ? '#e07070' : '#888',
              lineHeight: 1.6,
            }}>
              {overLimit ? (
                <>
                  <div style={{ fontWeight: 600, marginBottom: 6 }}>
                    ⚠ Files too large: {formatBytes(totalBytes)} (max {formatBytes(MAX_UPLOAD_BYTES)})
                  </div>
                  <div style={{ color: '#aaa', fontSize: 13 }}>
                    The receipts PDF is usually the heaviest. Compress it for free at{' '}
                    <a href="https://www.ilovepdf.com/compress_pdf" target="_blank" rel="noopener noreferrer"
                       style={{ color: '#c8a96e', textDecoration: 'underline' }}>iLovePDF</a>
                    {' '}or{' '}
                    <a href="https://smallpdf.com/compress-pdf" target="_blank" rel="noopener noreferrer"
                       style={{ color: '#c8a96e', textDecoration: 'underline' }}>SmallPDF</a>,
                    {' '}then re-upload it.
                  </div>
                </>
              ) : (
                <>
                  Total upload size: <span style={{ color: '#c8a96e' }}>{formatBytes(totalBytes)}</span>
                  {' '}of {formatBytes(MAX_UPLOAD_BYTES)}
                </>
              )}
            </div>
          )}

          <button
            onClick={handleGenerate}
            disabled={status === 'loading' || overLimit}
            style={{
              width: '100%', padding: '16px 32px',
              background: (status === 'loading' || overLimit) ? '#2a2a3a' : 'linear-gradient(135deg, #c8a96e, #8b6914)',
              color: (status === 'loading' || overLimit) ? '#666' : '#fff',
              border: 'none', borderRadius: 6, fontSize: 16,
              fontFamily: "'Georgia', serif", letterSpacing: '0.04em',
              cursor: (status === 'loading' || overLimit) ? 'not-allowed' : 'pointer', transition: 'all 0.2s',
            }}
          >
            {status === 'loading' ? '⏳ Generating...' : overLimit ? '✕ Files too large to upload' : '⬇ Generate & Download ZIP'}
          </button>

          {message && (
            <div style={{
              marginTop: 16, padding: '12px 16px', borderRadius: 6,
              background: status === 'error' ? 'rgba(200,60,60,0.1)' : 'rgba(60,180,100,0.1)',
              border: `1px solid ${status === 'error' ? '#8b2020' : '#2a6640'}`,
              color: status === 'error' ? '#e07070' : '#70c090', fontSize: 14,
            }}>{message}</div>
          )}
        </Section>

        <div style={{ marginTop: 48, padding: 24, background: '#0d0f14', borderRadius: 8, border: '1px solid #1e1e2e' }}>
          <div style={{ fontSize: 11, color: '#c8a96e', letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 16 }}>
            How It Works
          </div>
          <div style={{ fontSize: 13, color: '#888', lineHeight: 1.8 }}>
            Each Excel file becomes one invoice PDF with up to 3 pages:<br />
            <span style={{ color: '#c8a96e' }}>Page 1</span> — Invoice summary (from Excel Summary sheet)<br />
            <span style={{ color: '#c8a96e' }}>Page 2</span> — Work detail log (from Excel Work Detail sheet)<br />
            <span style={{ color: '#c8a96e' }}>Page 3+</span> — Receipt appended if matched by invoice number
          </div>
        </div>
      </div>
    </main>
  )
}

function Section({ number, title, children }: { number: string; title: string; children: React.ReactNode }) {
  return (
    <div style={{ marginBottom: 40 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{
          width: 28, height: 28, borderRadius: '50%',
          background: 'rgba(200,169,110,0.15)', border: '1px solid #c8a96e',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 13, color: '#c8a96e', fontWeight: 600, flexShrink: 0,
        }}>{number}</div>
        <h2 style={{ margin: 0, fontSize: 17, fontWeight: 500, color: '#e8e0d0', letterSpacing: '0.01em' }}>{title}</h2>
      </div>
      {children}
    </div>
  )
}

function SingleUploadBox({ label, accept, file, onChange, inputRef, icon }: {
  label: string
  accept: string
  file: File | null
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void
  inputRef: React.RefObject<HTMLInputElement>
  icon: string
}) {
  return (
    <div onClick={() => inputRef.current?.click()} style={{
      border: `2px dashed ${file ? '#c8a96e' : '#2a2a3a'}`,
      borderRadius: 8, padding: '28px 24px', textAlign: 'center', cursor: 'pointer',
      background: file ? 'rgba(200,169,110,0.04)' : '#0d0f14', transition: 'all 0.2s',
    }}>
      <input ref={inputRef} type="file" accept={accept} onChange={onChange} style={{ display: 'none' }} />
      {!file ? (
        <>
          <div style={{ fontSize: 28, marginBottom: 8 }}>{icon}</div>
          <div style={{ color: '#888', fontSize: 14 }}>{label}</div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 4 }}>PDF</div>
        </>
      ) : (
        <div style={{ textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span style={{ fontSize: 16 }}>✅</span>
            <span style={{ color: '#c8a96e', fontSize: 14 }}>{file.name}</span>
            <span style={{ color: '#555', fontSize: 12 }}>({(file.size / 1024).toFixed(1)} KB)</span>
          </div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 8 }}>Click to change</div>
        </div>
      )}
    </div>
  )
}

function MultiUploadBox({ label, accept, files, onChange, inputRef, icon }: {
  label: string
  accept: string
  files: File[]
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void
  inputRef: React.RefObject<HTMLInputElement>
  icon: string
}) {
  return (
    <div onClick={() => inputRef.current?.click()} style={{
      border: `2px dashed ${files.length > 0 ? '#c8a96e' : '#2a2a3a'}`,
      borderRadius: 8, padding: '28px 24px', textAlign: 'center', cursor: 'pointer',
      background: files.length > 0 ? 'rgba(200,169,110,0.04)' : '#0d0f14', transition: 'all 0.2s',
    }}>
      <input ref={inputRef} type="file" accept={accept} multiple onChange={onChange} style={{ display: 'none' }} />
      {files.length === 0 ? (
        <>
          <div style={{ fontSize: 28, marginBottom: 8 }}>{icon}</div>
          <div style={{ color: '#888', fontSize: 14 }}>{label}</div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 4 }}>{accept.toUpperCase().replace(/\./g, '').replace(/,/g, ' / ')}</div>
        </>
      ) : (
        <div style={{ textAlign: 'left' }}>
          {files.map(f => (
            <div key={f.name} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
              <span style={{ fontSize: 16 }}>✅</span>
              <span style={{ color: '#c8a96e', fontSize: 14 }}>{f.name}</span>
              <span style={{ color: '#555', fontSize: 12 }}>({(f.size / 1024).toFixed(1)} KB)</span>
            </div>
          ))}
          <div style={{ color: '#555', fontSize: 12, marginTop: 8 }}>Click to change</div>
        </div>
      )}
    </div>
  )
}
