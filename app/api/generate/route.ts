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

// ============================================================
// HELPERS
// ============================================================

function formatDate(val: unknown): string {
  if (!val) return ''
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
  // Allow letters, numbers, spaces, dashes, dots, hashes, commas, underscores.
  return name.replace(/[^a-zA-Z0-9_\-. #,]/g, '_').trim()
}

function fmtCurrency(val: unknown): string {
  const n = Number(val ?? 0)
  if (isNaN(n)) return String(val ?? '')
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n)
}

function fmtNumber(val: unknown, decimals = 2): string {
  const n = Number(val ?? 0)
  if (isNaN(n)) return String(val ?? '')
  return n.toFixed(decimals)
}

function stripLeasePrefix(s: string): string {
  return s.replace(/^lease\s+(no\.?\s*)?/i, '').trim()
}

function isEmpty(v: unknown): boolean {
  return v === '' || v === null || v === undefined
}

// Classify a column by its header to know how to format the values.
type ColKind = 'text' | 'days' | 'miles' | 'currency' | 'date'

function classifyHeader(header: string, isFirstCol = false): ColKind {
  const h = String(header).toLowerCase().trim()
  if (isFirstCol) return 'text'
  if (/^date$/.test(h)) return 'date'
  // Text indicators win first ("Miles Description" is text, not miles count)
  if (/broker|project|landman|prospect|legal|focus|description|complete/.test(h)) return 'text'
  // If the header mentions a rate/amount/total/etc., it's money even if "day" or "miles" is in the name
  const looksLikeMoney = /per|rate|amt|amount|services|total|copies|fee|mileage|labor|expense/.test(h)
  if (!looksLikeMoney) {
    if (/\bdays?\b|#\s*days/.test(h)) return 'days'
    if (/\bmiles\b/.test(h)) return 'miles'
  }
  return 'currency'
}

function fmtByKind(val: unknown, kind: ColKind): string {
  if (isEmpty(val)) return ''
  switch (kind) {
    case 'text': return String(val).trim()
    case 'days': return fmtNumber(val, 3)
    case 'miles': return String(Math.round(Number(val) || 0))
    case 'date': return formatDate(val)
    case 'currency': return fmtCurrency(val)
  }
}

// Render a row by classifying each cell according to its header.
function renderRow(row: unknown[], headers: string[]): string[] {
  return headers.map((h, i) => {
    const val = row[i]
    if (isEmpty(val)) return ''
    return fmtByKind(val, classifyHeader(h, i === 0))
  })
}

// ============================================================
// BILLING DATA LOOKUP
// ============================================================

type BillingEntry = {
  county: string
  type: string
  lease: string
  parcel: string
  unit: string
}

function parseBillingXlsx(buffer: Buffer): Map<string, BillingEntry> {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true })
  const sheetName = wb.SheetNames.find(n => n.toLowerCase() === 'parsed') || wb.SheetNames[0]
  const sheet = wb.Sheets[sheetName]
  if (!sheet) return new Map()

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as unknown[][]

  // Find header row.
  let headerIdx = -1
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (row.some(c => /invoice\s*(number|#|num)/i.test(String(c ?? '')))) {
      headerIdx = i
      break
    }
  }
  if (headerIdx < 0) return new Map()
  const headers = (rows[headerIdx] as unknown[]).map(h => String(h ?? '').toLowerCase().trim())

  const col = (...keywords: string[]): number => {
    for (const kw of keywords) {
      const i = headers.findIndex(h => h.includes(kw))
      if (i >= 0) return i
    }
    return -1
  }
  const invCol = col('invoice')
  const cntCol = col('county')
  const typeCol = col('type')
  const leaseCol = col('lease')
  const parcelCol = col('parcel', 'tpn', 'pid')
  const unitCol = col('unit')

  const map = new Map<string, BillingEntry>()
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i] as unknown[]
    const inv = String(r[invCol] ?? '').trim()
    if (!inv) continue
    map.set(inv, {
      county: cntCol >= 0 ? String(r[cntCol] ?? '').trim() : '',
      type: typeCol >= 0 ? String(r[typeCol] ?? '').trim() : '',
      lease: leaseCol >= 0 ? stripLeasePrefix(String(r[leaseCol] ?? '').trim()) : '',
      parcel: parcelCol >= 0 ? String(r[parcelCol] ?? '').trim() : '',
      unit: unitCol >= 0 ? String(r[unitCol] ?? '').trim() : '',
    })
  }
  return map
}

// ============================================================
// RECEIPT MATCHING
// ============================================================

function matchReceipt(invoiceNum: string, receiptFiles: { name: string; buffer: Buffer }[]): Buffer | null {
  for (const r of receiptFiles) {
    if (r.name.includes(invoiceNum)) return r.buffer
  }
  return null
}

// ============================================================
// PDF GENERATION
// ============================================================

async function buildInvoicePdf(
  excelBuffer: Buffer,
  billing: BillingEntry | undefined,
  receiptBuffer: Buffer | null,
  invoiceDateOverride: string,
): Promise<Buffer> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default

  const workbook = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
  const summarySheet = workbook.Sheets['Summary']
  const detailSheet = workbook.Sheets['Work Detail']
  if (!summarySheet) throw new Error('No Summary sheet found')

  const summaryRows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: '' }) as unknown[][]
  const invoiceNum = String((summaryRows[7] as unknown[])?.[5] ?? '').trim()
  const invoiceDate = invoiceDateOverride || String((summaryRows[7] as unknown[])?.[1] ?? '')
  const period = String((summaryRows[15] as unknown[])?.[1] ?? '')

  // From billing data, with sensible defaults.
  const lease = billing?.lease ? billing.lease : ''
  const parcel = billing?.parcel ? billing.parcel : ''
  const county = billing?.county || 'Washington'
  const unit = billing?.unit || ''
  const workType = billing?.type || 'Deed Search'

  // ----------------------------------------------------------
  // Find the broker-section header row (defaults to row 17).
  // ----------------------------------------------------------
  let brokerHeaderIdx = 17
  for (let i = 15; i < Math.min(summaryRows.length, 22); i++) {
    const r = summaryRows[i] as unknown[]
    if (r && /^broker$/i.test(String(r[0] ?? '').trim())) { brokerHeaderIdx = i; break }
  }
  const brokerHeadersRaw = ((summaryRows[brokerHeaderIdx] as unknown[]) || []).map(h => String(h ?? '').trim())
  while (brokerHeadersRaw.length > 0 && !brokerHeadersRaw[brokerHeadersRaw.length - 1]) brokerHeadersRaw.pop()
  const numCols = brokerHeadersRaw.length || 6

  // Identify the TOTAL column (case-insensitive match for "TOTAL").
  let totalColIdx = brokerHeadersRaw.findIndex(h => /^total$/i.test(h))
  if (totalColIdx < 0) totalColIdx = numCols - 1

  // Collect broker rows.
  const allBrokerRows: unknown[][] = []
  for (let i = brokerHeaderIdx + 1; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && row[0] && String(row[0]).trim()) allBrokerRows.push(row)
  }
  const dataRows = allBrokerRows.filter(r => String((r as unknown[])[0]).toLowerCase() !== 'totals')
  const totalsRow = allBrokerRows.find(r => String((r as unknown[])[0]).toLowerCase() === 'totals')

  // Add a "Project" column appended after TOTAL.
  const projectLabel = lease ? `Lease No. ${lease}` : ''
  const brokerHeaders = [...brokerHeadersRaw, 'Project']
  const projectColIdx = brokerHeaders.length - 1

  const brokerBody: string[][] = dataRows.map(r => {
    const rendered = renderRow(r as unknown[], brokerHeadersRaw)
    // Pad to numCols (in case the data row is shorter)
    while (rendered.length < numCols) rendered.push('')
    return [...rendered, projectLabel]
  })
  if (totalsRow) {
    const rendered = renderRow(totalsRow as unknown[], brokerHeadersRaw)
    while (rendered.length < numCols) rendered.push('')
    rendered[0] = 'Totals'
    brokerBody.push([...rendered, ''])
  }
  const brokerTotalsRowIndex = totalsRow ? brokerBody.length - 1 : -1

  // ----------------------------------------------------------
  // Work Detail dynamic parsing
  // ----------------------------------------------------------
  let detailHeaders: string[] = []
  let detailData: unknown[][] = []
  let detailTotalColIdx = -1
  if (detailSheet) {
    const detRows = XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][]
    // Header row is typically row index 1.
    let detHeaderIdx = 1
    for (let i = 0; i < Math.min(detRows.length, 5); i++) {
      const r = detRows[i] as unknown[]
      if (r && /^landman$/i.test(String(r[0] ?? '').trim())) { detHeaderIdx = i; break }
    }
    detailHeaders = ((detRows[detHeaderIdx] as unknown[]) || []).map(h => String(h ?? '').trim())
    while (detailHeaders.length > 0 && !detailHeaders[detailHeaders.length - 1]) detailHeaders.pop()
    detailTotalColIdx = detailHeaders.findIndex(h => /^total$/i.test(h))
    detailData = detRows.slice(detHeaderIdx + 1).filter(r => {
      const row = r as unknown[]
      return row && row[0] && String(row[0]).trim() !== ''
    })
  }

  // ----------------------------------------------------------
  // Colors and constants
  // ----------------------------------------------------------
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

  const leftLabel = (label: string, value: string) => {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
    doc.text(label, lx, ly)
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
    doc.text(value, lx + 45, ly)
    ly += 13
  }
  const rightLabel = (label: string, value: string) => {
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

  rightLabel('Lease No.:', lease)
  rightLabel('PID:', parcel)
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

  // ----------------------------------------------------------
  // Dynamic broker-summary column widths
  // ----------------------------------------------------------
  const pageContentWidth = 532 // 612 - 80 margins
  const fixedBrokerW = 75
  const fixedProjectW = 100
  const totalColW = 62 // give TOTAL a touch more
  const middleCount = brokerHeaders.length - 2 // exclude Broker + Project
  const middleAvail = pageContentWidth - fixedBrokerW - fixedProjectW - totalColW
  const middleColW = Math.max(40, Math.floor(middleAvail / Math.max(1, middleCount - 1)))

  const brokerColStyles: Record<number, { cellWidth: number; halign?: 'left' | 'center' | 'right'; overflow?: 'linebreak' }> = {}
  brokerHeaders.forEach((h, i) => {
    if (i === 0) {
      brokerColStyles[i] = { cellWidth: fixedBrokerW, halign: 'left' }
    } else if (i === projectColIdx) {
      brokerColStyles[i] = { cellWidth: fixedProjectW, overflow: 'linebreak' }
    } else if (i === totalColIdx) {
      brokerColStyles[i] = { cellWidth: totalColW, halign: 'right' }
    } else {
      const kind = classifyHeader(h, false)
      brokerColStyles[i] = {
        cellWidth: middleColW,
        halign: kind === 'days' || kind === 'miles' ? 'center' : 'right',
      }
    }
  })

  autoTable(doc, {
    startY: tableStartY,
    head: [brokerHeaders],
    body: brokerBody,
    theme: 'grid',
    styles: {
      font: 'helvetica', fontSize: 8, textColor: black,
      fillColor: white, cellPadding: 4, lineColor: black, lineWidth: 0.3,
      overflow: 'linebreak',
    },
    headStyles: { textColor: black, fillColor: headerBg, fontStyle: 'bold', halign: 'center' },
    columnStyles: brokerColStyles,
    didParseCell: function(data) {
      if (data.section === 'body' && data.row.index === brokerTotalsRowIndex) {
        data.cell.styles.fontStyle = 'bold'
        if (data.column.index === totalColIdx) {
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
  if (detailHeaders.length > 0 && detailData.length > 0) {
    doc.addPage()
    doc.setFillColor(...white)
    doc.rect(0, 0, 612, 792, 'F')

    doc.setFont('helvetica', 'bold')
    doc.setFontSize(11)
    doc.setTextColor(...black)
    doc.text('Work Detail', 40, 45)
    doc.setLineWidth(0.5)
    doc.line(40, 52, 572, 52)

    // Compute totals row for the work detail by summing numeric columns.
    // Skip per-unit rate columns: "Amt. Per Day", "Dayrate Amount", "Hourly Rate", etc.
    // — these are rates, not amounts to be summed.
    // DO sum calculated amounts even if header mentions "/mile" (e.g. "Mileage 0.7250/mile").
    const isRateColumn = (h: string) =>
      /\brate\b|dayrate|per\s+(day|diem|hour|hr|mile)|amt\.?\s*per|amount\s*per/i.test(h)
    const totalsRowDetail: string[] = new Array(detailHeaders.length).fill('')
    totalsRowDetail[0] = 'Totals'
    detailHeaders.forEach((h, i) => {
      if (i === 0) return
      if (isRateColumn(h)) return
      const kind = classifyHeader(h, false)
      if (kind === 'currency' || kind === 'days' || kind === 'miles') {
        let sum = 0
        let any = false
        for (const r of detailData) {
          const v = (r as unknown[])[i]
          const n = Number(v)
          if (!isNaN(n) && v !== '' && v !== null && v !== undefined) {
            sum += n
            any = true
          }
        }
        if (any) totalsRowDetail[i] = fmtByKind(sum, kind)
      }
    })

    const detailBody: string[][] = detailData.map(r => renderRow(r as unknown[], detailHeaders))
    detailBody.push(totalsRowDetail)
    const detailTotalsIndex = detailBody.length - 1

    // Dynamic Work Detail widths.
    // Use weight-based proportional layout, then scale to fit page.
    const colWeight = (h: string): number => {
      const lo = h.toLowerCase()
      if (/description|complete/.test(lo) && !/miles\s+description/.test(lo)) return 14
      if (/miles\s+description/.test(lo)) return 8
      if (/landman|prospect/.test(lo)) return 7
      if (/^date$/.test(lo)) return 6
      if (/legal/.test(lo)) return 8
      if (/focus|lease/.test(lo)) return 7
      if (/\bdays\b/.test(lo)) return 4
      if (/\bmiles\b/.test(lo) && !/mileage|description/.test(lo)) return 4
      if (/copies/.test(lo)) return 6
      if (/total\b/.test(lo)) return 6
      if (/mileage|rate|amount|services|dayrate/.test(lo)) return 6
      return 6
    }
    const weights = detailHeaders.map(colWeight)
    const totalWeight = weights.reduce((a, b) => a + b, 0) || 1
    const detailColStyles: Record<number, { cellWidth: number; halign?: 'left' | 'center' | 'right'; overflow?: 'linebreak' }> = {}
    detailHeaders.forEach((h, i) => {
      const w = Math.max(28, Math.floor((weights[i] / totalWeight) * pageContentWidth))
      const kind = classifyHeader(h, i === 0)
      const isText = kind === 'text' || kind === 'date'
      detailColStyles[i] = {
        cellWidth: w,
        halign: kind === 'currency' ? 'right' : kind === 'days' || kind === 'miles' ? 'center' : 'left',
        ...(isText ? { overflow: 'linebreak' } : {}),
      }
    })

    autoTable(doc, {
      startY: 60,
      head: [detailHeaders],
      body: detailBody,
      theme: 'grid',
      styles: {
        font: 'helvetica', fontSize: 7, textColor: black,
        fillColor: white, cellPadding: 3, lineColor: black, lineWidth: 0.3,
        overflow: 'linebreak', valign: 'top',
      },
      headStyles: {
        textColor: black, fillColor: headerBg, fontStyle: 'bold',
        halign: 'center', overflow: 'linebreak', valign: 'middle',
      },
      columnStyles: detailColStyles,
      didParseCell: function(data) {
        if (data.section === 'body' && data.row.index === detailTotalsIndex) {
          data.cell.styles.fontStyle = 'bold'
          if (data.column.index === detailTotalColIdx) {
            data.cell.styles.fillColor = totalsBg
          }
        }
      },
      margin: { left: 40, right: 40 },
    })
  }

  // ============ APPEND RECEIPT ============
  try {
    const { PDFDocument } = await import('pdf-lib')
    const mainPdfBytes = doc.output('arraybuffer')
    const mainPdf = await PDFDocument.load(mainPdfBytes)

    if (receiptBuffer) {
      const receiptPdf = await PDFDocument.load(receiptBuffer)
      const receiptPages = await mainPdf.copyPages(receiptPdf, receiptPdf.getPageIndices())
      receiptPages.forEach(p => mainPdf.addPage(p))
    }

    const finalBytes = await mainPdf.save()
    return Buffer.from(finalBytes)
  } catch {
    return Buffer.from(doc.output('arraybuffer'))
  }
}

// ============================================================
// POST HANDLER
// ============================================================

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const excelFiles = formData.getAll('excel') as File[]
    const receiptFiles = formData.getAll('receipts') as File[]
    const billingFile = formData.get('billing') as File | null
    const rawDate = (formData.get('invoiceDate') as string) || ''
    const invoiceDateOverride = normalizeInvoiceDate(rawDate)
    const outputFormat = ((formData.get('outputFormat') as string) || 'pdf').toLowerCase()
    const wantsPdf = outputFormat === 'pdf' || outputFormat === 'both'
    const wantsXlsx = outputFormat === 'excel' || outputFormat === 'both'

    if (!excelFiles.length) return NextResponse.json({ error: 'No Excel files uploaded' }, { status: 400 })
    if (!billingFile) return NextResponse.json({ error: 'No billing data spreadsheet uploaded' }, { status: 400 })

    const billingBuffer = Buffer.from(await billingFile.arrayBuffer())
    const billingMap = parseBillingXlsx(billingBuffer)

    const receiptData: { name: string; buffer: Buffer }[] = []
    for (const r of receiptFiles) {
      receiptData.push({ name: r.name, buffer: Buffer.from(await r.arrayBuffer()) })
    }

    const zip = new JSZip()

    for (const excelFile of excelFiles) {
      const excelBuffer = Buffer.from(await excelFile.arrayBuffer())
      const wb = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
      const ws = wb.Sheets['Summary']
      if (!ws) continue
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]
      const invoiceNum = String((rows[7] as unknown[])?.[5] ?? '').trim() ||
        (excelFile.name.match(/(\d{5,})/)?.[1] ?? 'UNKNOWN')

      const billing = billingMap.get(invoiceNum)

      // Filename components.
      const fnUnit = billing?.unit || 'Unknown'
      const fnType = billing?.type || 'Deed Search'
      let fnParcel = billing?.parcel || ''
      if (!fnParcel) {
        // Fallback: try to grab the first PID-looking thing from the Work Detail sheet.
        const detailWs = wb.Sheets['Work Detail']
        if (detailWs) {
          const detRows = XLSX.utils.sheet_to_json(detailWs, { header: 1, defval: '' }) as unknown[][]
          for (const r of detRows.slice(1)) {
            const cell = String((r as unknown[])[3] ?? '').trim()
            const m = cell.match(/\d{3}-\d{3}-\d{2}-\d{2}-\d{4}-\d{2}/)
            if (m) { fnParcel = m[0]; break }
          }
        }
        if (!fnParcel) fnParcel = 'Unknown'
      }

      const outputName = sanitize(`#${invoiceNum} - ${fnUnit} - ${fnParcel} - ${fnType}`)

      const matchedReceipt = matchReceipt(invoiceNum, receiptData)

      try {
        if (wantsPdf) {
          const pdfBuffer = await buildInvoicePdf(excelBuffer, billing, matchedReceipt, invoiceDateOverride)
          zip.file(`${outputName}.pdf`, pdfBuffer)
        }
        if (wantsXlsx) {
          // Pass-through the source Excel renamed.
          zip.file(`${outputName}.xlsx`, excelBuffer)
        }
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
