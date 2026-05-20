import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'

const BILL_TO = {
  company: 'Range Resources - Appalachia, LLC',
  attn: 'Attn: Laura Schimmel',
  address: '3000 Town Center Blvd.',
  city: 'Canonsburg, Pennsylvania 15317',
}

const BOP = {
  name: 'BOP Abstract, LLC',
  address: '2547 Washington Rd. Bldg. 700, Ste. 720',
  city: 'Pittsburgh, PA, 15241',
  phone: '724-747-1594',
}

function formatDate(val: unknown): string {
  if (!val) return new Date().toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  if (val instanceof Date) return val.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  const d = new Date(String(val))
  if (!isNaN(d.getTime())) return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  return String(val)
}

function normalizeInvoiceDate(input: string): string {
  const d = new Date(input)
  if (!isNaN(d.getTime())) {
    return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
  }
  return input
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_\-. #]/g, '_').trim()
}

function cleanPid(val: string): string {
  return val.replace(/^[-\s]+/, '').trim()
}

function fmtCurrency(val: unknown): string {
  const n = Number(val ?? 0)
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n)
}

// Extracts every PID-looking token from a string.
function extractAllPids(s: string): string[] {
  if (!s) return []
  const matches = s.match(/\d{3}-\d{3}-\d{2}-\d{2}-\d{4}-\d{2}/g)
  return matches ? matches.map(cleanPid) : []
}

// Fix names that got broken by PDF text extraction (e.g. "S tevenson" -> "Stevenson").
function fixBrokenName(s: string): string {
  return s
    .replace(/\b([A-Z])\s+([a-z]{2,})/g, '$1$2')
    .replace(/\s+/g, ' ')
    .trim()
}

type EmailEntry = {
  pids: string[]
  unit: string
  county: string
  type: string
  lease: string
}

function parseEmailPdf(text: string): {
  byLease: Map<string, EmailEntry>
  byPid: Map<string, EmailEntry>
} {
  const byLease = new Map<string, EmailEntry>()
  const byPid = new Map<string, EmailEntry>()

  // Flatten the document so paragraph wraps don't break our patterns.
  const flat = text.replace(/\s*\n\s*/g, ' ').replace(/\s+/g, ' ')

  // Find every "Please follow the Dropbox link..." marker location.
  const markerRe = /Please follow the Dropbox link below for the completed abstracts on/gi
  const starts: number[] = []
  let m: RegExpExecArray | null
  while ((m = markerRe.exec(flat)) !== null) {
    starts.push(m.index + m[0].length)
  }

  for (let i = 0; i < starts.length; i++) {
    const segStart = starts[i]
    // Cap the segment at the NEXT marker, OR at email-boundary markers
    // (subject line of next email, "Thank You" signoff, "Note:" annotation).
    const nextMarker = i + 1 < starts.length ? starts[i + 1] : flat.length
    const stopCandidates = [
      nextMarker,
      flat.indexOf('From:', segStart),
      flat.indexOf('Subject:', segStart),
      flat.indexOf('Note:', segStart),
      flat.toLowerCase().indexOf('thank you,', segStart),
    ].filter(x => x > segStart)
    const segEnd = stopCandidates.length > 0 ? Math.min(...stopCandidates) : flat.length

    const seg = flat.slice(segStart, segEnd)

    // Header: "[County] County - [Work Type] - (Lease [Lease#] | Unleased) - ..."
    const headerRe = /([A-Z][a-zA-Z]+)\s+County\s*[-–]\s*([^-–]+?)\s*[-–]\s*(?:Lease\s+([A-Za-z0-9\-]+)|Unleased)\b/
    const headerMatch = seg.match(headerRe)
    if (!headerMatch) continue

    const county = headerMatch[1].trim()
    const type = headerMatch[2].trim()
    const lease = (headerMatch[3] ?? '').trim()

    // Work in the slice AFTER the header so we don't pick up PIDs from a
    // preceding subject line or signature.
    const headerIdx = seg.indexOf(headerMatch[0])
    const afterHeader = seg.slice(headerIdx + headerMatch[0].length)

    // PIDs ONLY live between the header and the "(Unit Name)" parens
    // (or the first dropbox URL, whichever comes first). This prevents
    // counting PIDs from the RESTATED block / URL line.
    const parenIdx = afterHeader.search(/\([^)]+\)/)
    const urlIdx = afterHeader.search(/https?:\/\//)
    const stops = [parenIdx, urlIdx].filter(x => x >= 0)
    const pidScanEnd = stops.length > 0 ? Math.min(...stops) : afterHeader.length
    const pidScanRegion = afterHeader.slice(0, pidScanEnd)

    // Dedupe in case the same PID appears more than once in the same paragraph.
    const pids = Array.from(new Set(extractAllPids(pidScanRegion)))

    // Unit = first non-empty parenthesized phrase, skipping URL-looking parens.
    let unit = ''
    const parenMatches = afterHeader.match(/\(([^)]+)\)/g)
    if (parenMatches) {
      for (const p of parenMatches) {
        const inner = p.slice(1, -1).trim()
        if (!inner) continue
        if (/https?:|dropbox|\.com|\.pdf/i.test(inner)) continue
        unit = fixBrokenName(inner)
        break
      }
    }

    const entry: EmailEntry = { pids, unit, county, type, lease }

    if (lease) byLease.set(lease, entry)
    for (const p of pids) {
      if (!byPid.has(p)) byPid.set(p, entry)
    }
  }

  return { byLease, byPid }
}

function matchReceipt(invoiceNum: string, receiptFiles: { name: string; buffer: Buffer }[]): Buffer | null {
  for (const r of receiptFiles) {
    if (r.name.includes(invoiceNum)) return r.buffer
  }
  return null
}

function lookupEmailInfo(
  excelLease: string,
  excelPidCell: string,
  byLease: Map<string, EmailEntry>,
  byPid: Map<string, EmailEntry>,
): EmailEntry | undefined {
  if (excelLease) {
    const hit = byLease.get(excelLease.trim())
    if (hit) return hit
    const norm = excelLease.trim().toLowerCase()
    for (const [k, v] of byLease) {
      if (k.toLowerCase() === norm) return v
    }
  }
  for (const p of extractAllPids(excelPidCell)) {
    const hit = byPid.get(p)
    if (hit) return hit
  }
  return undefined
}

async function buildInvoicePdf(
  excelBuffer: Buffer,
  emailText: string,
  matchedReceiptBuffer: Buffer | null,
  invoiceDateOverride: string,
): Promise<Buffer> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default

  const workbook = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
  const summarySheet = workbook.Sheets['Summary']
  const detailSheet = workbook.Sheets['Work Detail']

  if (!summarySheet) throw new Error('No Summary sheet found')

  const summaryRows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: '' }) as unknown[][]
  const invoiceNum = String((summaryRows[7] as unknown[])?.[5] ?? '')
  const invoiceDate = invoiceDateOverride || String((summaryRows[7] as unknown[])?.[1] ?? '')
  const period = String((summaryRows[15] as unknown[])?.[1] ?? '')

  const headerRow = summaryRows[17] as string[]
  const hasCopiesCol = headerRow && headerRow.some(h => String(h).toLowerCase().includes('cop'))

  const brokerRows: unknown[][] = []
  for (let i = 18; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && row[0] && String(row[0]).trim()) brokerRows.push(row)
  }

  const brokerDataRows = brokerRows.filter(r => String((r as unknown[])[0]).toLowerCase() !== 'totals')
  const brokerTotalsRow = brokerRows.find(r => String((r as unknown[])[0]).toLowerCase() === 'totals')

  const detailRows = detailSheet
    ? (XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][]).slice(2)
    : []

  let leaseNo = ''
  let pidCell = ''
  let pid = ''
  let unit = ''
  let county = 'Washington'
  let workType = 'Deed Search'

  for (const row of detailRows) {
    const r = row as unknown[]
    if (r[4] && String(r[4]).trim()) {
      leaseNo = String(r[4]).trim()
      pidCell = cleanPid(String(r[3] ?? '').trim())
      pid = extractAllPids(pidCell)[0] || pidCell
      break
    }
  }

  const { byLease, byPid } = parseEmailPdf(emailText)
  const emailInfo = lookupEmailInfo(leaseNo, pidCell, byLease, byPid)

  if (emailInfo) {
    if (emailInfo.unit) unit = emailInfo.unit
    if (emailInfo.county) county = emailInfo.county
    if (emailInfo.type) workType = emailInfo.type
  }

  const black = [0, 0, 0] as [number, number, number]
  const red = [255, 0, 0] as [number, number, number]
  const headerBg = [242, 220, 219] as [number, number, number]
  const totalsBg = [255, 255, 0] as [number, number, number]
  const white = [255, 255, 255] as [number, number, number]
  const lightGray = [100, 100, 100] as [number, number, number]

  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'letter' })

  // ============ PAGE 1: SUMMARY ============
  doc.setFillColor(...white)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.setTextColor(...black)
  doc.text(BOP.name, 306, 45, { align: 'center' })
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.text(BOP.address, 306, 57, { align: 'center' })
  doc.text(BOP.city, 306, 69, { align: 'center' })
  doc.text(BOP.phone, 306, 81, { align: 'center' })

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(13)
  doc.text('INVOICE', 306, 105, { align: 'center' })

  doc.setDrawColor(...black)
  doc.setLineWidth(0.5)
  doc.line(40, 112, 572, 112)

  doc.setFontSize(9)
  const lx = 40
  const rx = 320
  let ly = 128
  let ry = 128

  function leftLabel(label: string, value: string) {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
    doc.text(label, lx, ly)
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
    doc.text(value, lx + 45, ly)
    ly += 13
  }

  function rightLabel(label: string, value: string) {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
    doc.text(label, rx, ry)
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
    doc.text(value, rx + 65, ry)
    ry += 13
  }

  leftLabel('Date:', invoiceDate)
  rightLabel('Invoice #:', invoiceNum)
  ly += 4; ry += 4

  doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
  doc.text('Bill To:', lx, ly)
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...black)
  doc.text(BILL_TO.company, lx + 45, ly); ly += 13
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
  doc.text(BILL_TO.attn, lx + 45, ly); ly += 13
  doc.text(BILL_TO.address, lx + 45, ly); ly += 13
  doc.text(BILL_TO.city, lx + 45, ly); ly += 13

  rightLabel('Lease No.:', leaseNo)
  rightLabel('PID:', pid)
  rightLabel('County:', county)
  rightLabel('Unit:', unit)
  rightLabel('Type:', workType)

  ly += 8
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
  doc.text('Period:', lx, ly)
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
  doc.text(period, lx + 45, ly)

  doc.setTextColor(...red)
  doc.setFont('helvetica', 'bold')
  doc.text('DUE UPON RECEIPT', 572, ly, { align: 'right' })
  doc.setTextColor(...black)

  const tableStartY = Math.max(ly + 20, ry + 20)

  let tableHead: string[][]
  let tableBody: string[][]
  const totalCol = hasCopiesCol ? 5 : 4

  if (hasCopiesCol) {
    tableHead = [['Broker', '# Days', 'Amt. Per Day', 'Total Prof. Services', 'Copies', 'TOTAL', 'Project']]
    tableBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        Number(row[1] ?? 0).toFixed(3),
        fmtCurrency(row[2]),
        fmtCurrency(row[3]),
        fmtCurrency(row[4]),
        fmtCurrency(row[5]),
        `Lease No. ${leaseNo}`,
      ]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      tableBody.push(['Totals', Number(t[1] ?? 0).toFixed(3), '', fmtCurrency(t[3]), fmtCurrency(t[4]), fmtCurrency(t[5]), ''])
    }
  } else {
    tableHead = [['Broker', '# Days', 'Amt. Per Day', 'Total Prof. Services', 'TOTAL', 'Project']]
    tableBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        Number(row[1] ?? 0).toFixed(3),
        fmtCurrency(row[2]),
        fmtCurrency(row[3]),
        fmtCurrency(row[4]),
        `Lease No. ${leaseNo}`,
      ]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      tableBody.push(['Totals', Number(t[1] ?? 0).toFixed(3), '', fmtCurrency(t[3]), fmtCurrency(t[4]), ''])
    }
  }

  const totalsRowIndex = tableBody.length - 1

  autoTable(doc, {
    startY: tableStartY,
    head: tableHead,
    body: tableBody,
    theme: 'grid',
    styles: {
      font: 'helvetica', fontSize: 8, textColor: black,
      fillColor: white, cellPadding: 4, lineColor: black, lineWidth: 0.3,
      overflow: 'linebreak',
    },
    headStyles: { textColor: black, fillColor: headerBg, fontStyle: 'bold', halign: 'center' },
    columnStyles: {
      0: { cellWidth: 75 },
      1: { halign: 'center', cellWidth: 42 },
      2: { halign: 'right', cellWidth: 62 },
      3: { halign: 'right', cellWidth: 82 },
      4: { halign: 'right', cellWidth: 62 },
      5: { halign: 'right', cellWidth: 62 },
      6: { cellWidth: 105, overflow: 'linebreak' },
    },
    didParseCell: function(data) {
      if (data.section === 'body' && data.row.index === totalsRowIndex) {
        data.cell.styles.fontStyle = 'bold'
        if (data.column.index === totalCol) {
          data.cell.styles.fillColor = totalsBg
        }
      }
    },
    margin: { left: 40, right: 40 },
  })

  const finalY = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20
  doc.setFontSize(8)
  doc.setTextColor(...lightGray)
  doc.setFont('helvetica', 'italic')
  doc.text('Please contact our accounting department with any questions regarding invoices', 306, finalY, { align: 'center' })

  // ============ PAGE 2: WORK DETAIL ============
  doc.addPage()
  doc.setFillColor(...white)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.setTextColor(...black)
  doc.text('Work Detail', 40, 45)
  doc.setLineWidth(0.5)
  doc.line(40, 52, 572, 52)

  const hasDetailCopies = detailSheet
    ? (XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][])[1]?.some(
        (h: unknown) => String(h).toLowerCase().includes('cop')
      )
    : false

  const detailDataRows = detailRows.filter(r => {
    const row = r as unknown[]
    return row[0] && String(row[0]).trim() !== ''
  })

  let totalDays = 0
  let totalLaborTotal = 0
  let totalCopies = 0
  let totalTotal = 0

  detailDataRows.forEach(r => {
    const row = r as unknown[]
    if (hasDetailCopies) {
      totalDays += Number(row[5] ?? 0)
      totalLaborTotal += Number(row[7] ?? 0)
      totalCopies += Number(row[8] ?? 0)
      totalTotal += Number(row[9] ?? 0)
    } else {
      totalDays += Number(row[5] ?? 0)
      totalLaborTotal += Number(row[7] ?? 0)
      totalTotal += Number(row[8] ?? 0)
    }
  })

  let detailHead: string[][]
  let detailBody: string[][]
  const detailTotalCol = hasDetailCopies ? 8 : 7

  if (hasDetailCopies) {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Lease No.', 'Days', 'Labor\nTotal', 'Copies', 'Total', 'Description']]
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        formatDate(row[1]),
        String(row[2] ?? ''),
        cleanPid(String(row[3] ?? '')),
        String(row[4] ?? ''),
        Number(row[5] ?? 0).toFixed(2),
        fmtCurrency(row[7]),
        fmtCurrency(row[8]),
        fmtCurrency(row[9]),
        String(row[10] ?? ''),
      ]
    })
    detailBody.push(['Totals', '', '', '', '', totalDays.toFixed(2), fmtCurrency(totalLaborTotal), fmtCurrency(totalCopies), fmtCurrency(totalTotal), ''])
  } else {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Lease No.', 'Days', 'Labor\nTotal', 'Total', 'Description']]
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        formatDate(row[1]),
        String(row[2] ?? ''),
        cleanPid(String(row[3] ?? '')),
        String(row[4] ?? ''),
        Number(row[5] ?? 0).toFixed(2),
        fmtCurrency(row[7]),
        fmtCurrency(row[8]),
        String(row[9] ?? ''),
      ]
    })
    detailBody.push(['Totals', '', '', '', '', totalDays.toFixed(2), fmtCurrency(totalLaborTotal), fmtCurrency(totalTotal), ''])
  }

  const detailTotalsIndex = detailBody.length - 1

  autoTable(doc, {
    startY: 60,
    head: detailHead,
    body: detailBody,
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 7,
      textColor: black,
      fillColor: white,
      cellPadding: 3,
      lineColor: black,
      lineWidth: 0.3,
      overflow: 'linebreak',
      valign: 'top',
    },
    headStyles: {
      textColor: black,
      fillColor: headerBg,
      fontStyle: 'bold',
      halign: 'center',
      overflow: 'linebreak',
      valign: 'middle',
    },
    columnStyles: hasDetailCopies ? {
      0: { cellWidth: 55, overflow: 'linebreak' },
      1: { cellWidth: 45, overflow: 'linebreak' },
      2: { cellWidth: 55, overflow: 'linebreak' },
      3: { cellWidth: 65, overflow: 'linebreak' },
      4: { cellWidth: 55, overflow: 'linebreak' },
      5: { cellWidth: 28, halign: 'center' },
      6: { cellWidth: 48, halign: 'right' },
      7: { cellWidth: 48, halign: 'right' },
      8: { cellWidth: 48, halign: 'right' },
      9: { cellWidth: 85, overflow: 'linebreak' },
    } : {
      0: { cellWidth: 55, overflow: 'linebreak' },
      1: { cellWidth: 45, overflow: 'linebreak' },
      2: { cellWidth: 55, overflow: 'linebreak' },
      3: { cellWidth: 65, overflow: 'linebreak' },
      4: { cellWidth: 55, overflow: 'linebreak' },
      5: { cellWidth: 28, halign: 'center' },
      6: { cellWidth: 55, halign: 'right' },
      7: { cellWidth: 55, halign: 'right' },
      8: { cellWidth: 119, overflow: 'linebreak' },
    },
    didParseCell: function(data) {
      if (data.section === 'body' && data.row.index === detailTotalsIndex) {
        data.cell.styles.fontStyle = 'bold'
        if (data.column.index === detailTotalCol) {
          data.cell.styles.fillColor = totalsBg
        }
      }
    },
    margin: { left: 40, right: 40 },
  })

  // ============ APPEND RECEIPT (if matched) ============
  try {
    const { PDFDocument } = await import('pdf-lib')
    const mainPdfBytes = doc.output('arraybuffer')
    const mainPdf = await PDFDocument.load(mainPdfBytes)

    if (matchedReceiptBuffer) {
      const receiptPdf = await PDFDocument.load(matchedReceiptBuffer)
      const receiptPages = await mainPdf.copyPages(receiptPdf, receiptPdf.getPageIndices())
      receiptPages.forEach(p => mainPdf.addPage(p))
    }

    const finalBytes = await mainPdf.save()
    return Buffer.from(finalBytes)
  } catch {
    return Buffer.from(doc.output('arraybuffer'))
  }
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const excelFiles = formData.getAll('excel') as File[]
    const receiptFiles = formData.getAll('receipts') as File[]
    const emailFile = formData.get('email') as File
    const rawDate = (formData.get('invoiceDate') as string) || ''
    const invoiceDateOverride = normalizeInvoiceDate(rawDate)

    if (!excelFiles.length) return NextResponse.json({ error: 'No Excel files uploaded' }, { status: 400 })
    if (!emailFile) return NextResponse.json({ error: 'No email PDF uploaded' }, { status: 400 })

    const emailBuffer = Buffer.from(await emailFile.arrayBuffer())
    let emailText = ''
    try {
      const pdfParse = (await import('pdf-parse')).default
      const parsed = await pdfParse(emailBuffer)
      emailText = parsed.text
    } catch {
      emailText = ''
    }

    const receiptData: { name: string; buffer: Buffer }[] = []
    for (const r of receiptFiles) {
      receiptData.push({ name: r.name, buffer: Buffer.from(await r.arrayBuffer()) })
    }

    const { byLease, byPid } = parseEmailPdf(emailText)

    const zip = new JSZip()

    for (const excelFile of excelFiles) {
      const excelBuffer = Buffer.from(await excelFile.arrayBuffer())
      const wb = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
      const ws = wb.Sheets['Summary']
      const detailWs = wb.Sheets['Work Detail']
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]
      const invoiceNum = String((rows[7] as unknown[])?.[5] ?? '').trim() ||
        (excelFile.name.match(/(\d{5,})/)?.[1] ?? 'UNKNOWN')

      const detailRows = detailWs
        ? (XLSX.utils.sheet_to_json(detailWs, { header: 1, defval: '' }) as unknown[][]).slice(2)
        : []

      let filenamePidCell = ''
      let filenameLeaseNo = ''
      for (const row of detailRows) {
        const r = row as unknown[]
        if (r[4] && String(r[4]).trim()) {
          filenameLeaseNo = String(r[4]).trim()
          filenamePidCell = cleanPid(String(r[3] ?? '').trim())
          break
        }
      }

      const emailInfo = lookupEmailInfo(filenameLeaseNo, filenamePidCell, byLease, byPid)
      const filenameUnit = emailInfo?.unit || ''
      const filenameType = emailInfo?.type || 'Deed Search'

      // Filename PID: take the FIRST pid from the email block (authoritative).
      // Append " et al" ONLY if the email block lists more than one PID.
      const emailPids = emailInfo?.pids || []
      const excelPids = extractAllPids(filenamePidCell)
      const firstPid = emailPids[0] || excelPids[0] || ''
      const hasMultiple = emailPids.length > 1
      const filenamePid = firstPid
        ? (hasMultiple ? `${firstPid} et al` : firstPid)
        : (filenamePidCell || 'Unknown')

      const nameParts = [
        `#${invoiceNum}`,
        filenameUnit || 'Unknown',
        filenamePid,
        filenameType,
      ]
      const outputName = sanitize(nameParts.join(' - '))

      const matchedReceipt = matchReceipt(invoiceNum, receiptData)

      try {
        const pdfBuffer = await buildInvoicePdf(excelBuffer, emailText, matchedReceipt, invoiceDateOverride)
        zip.file(`${outputName}.pdf`, pdfBuffer)
      } catch (err) {
        return NextResponse.json({ error: `Failed for invoice ${invoiceNum}: ${String(err)}` }, { status: 500 })
      }
    }

    const zipBuffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
    return new NextResponse(zipBuffer, {
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': 'attachment; filename="range-invoices.zip"',
      },
    })
  } catch (err) {
    console.error(err)
    return NextResponse.json({ error: String(err) }, { status: 500 })
  }
}
