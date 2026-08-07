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

  // Find the sheet that actually has an "Invoice Number" header row.
  // The billing spreadsheet may have multiple sheets (e.g. "Email List" + "Invoice Details");
  // we want the one with the lookup table, regardless of its name.
  let rows: unknown[][] = []
  let headerIdx = -1
  for (const name of wb.SheetNames) {
    const candidate = wb.Sheets[name]
    if (!candidate) continue
    const candidateRows = XLSX.utils.sheet_to_json(candidate, { header: 1, defval: '' }) as unknown[][]
    for (let i = 0; i < Math.min(candidateRows.length, 5); i++) {
      const r = candidateRows[i] as unknown[]
      if (r && r.some(c => /invoice\s*(number|#|num)/i.test(String(c ?? '')))) {
        rows = candidateRows
        headerIdx = i
        break
      }
    }
    if (headerIdx >= 0) break
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
// EXCEL INVOICE GENERATION (mirrors the PDF layout, fully styled)
// ============================================================

// Colors matching the PDF
const XLSX_COLOR_PINK = 'FFF2DCDB'    // Table header background
const XLSX_COLOR_YELLOW = 'FFFFFF00'  // Totals highlight
const XLSX_COLOR_RED = 'FFFF0000'     // DUE UPON RECEIPT
const XLSX_COLOR_GRAY = 'FF646464'    // Label text
const XLSX_COLOR_BLACK = 'FF000000'
const XLSX_CURRENCY_FMT = '"$"#,##0.00;("$"#,##0.00)'
const XLSX_DAYS_FMT = '0.000'

async function buildInvoiceXlsx(
  excelBuffer: Buffer,
  billing: BillingEntry | undefined,
  invoiceDateOverride: string,
): Promise<Buffer> {
  const ExcelJS = (await import('exceljs')).default
  const workbook = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
  const summarySheet = workbook.Sheets['Summary']
  const detailSheet = workbook.Sheets['Work Detail']
  if (!summarySheet) throw new Error('No Summary sheet found')

  const summaryRows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: '' }) as unknown[][]
  const invoiceNum = String((summaryRows[7] as unknown[])?.[5] ?? '').trim()
  const invoiceDate = invoiceDateOverride || String((summaryRows[7] as unknown[])?.[1] ?? '')
  const period = String((summaryRows[15] as unknown[])?.[1] ?? '')

  const lease = billing?.lease ?? ''
  const parcel = billing?.parcel ?? ''
  const county = billing?.county || 'Washington'
  const unit = billing?.unit ?? ''
  const workType = billing?.type || 'Deed Search'

  // ---- Parse broker table ----
  let brokerHeaderIdx = 17
  for (let i = 15; i < Math.min(summaryRows.length, 22); i++) {
    const r = summaryRows[i] as unknown[]
    if (r && /^broker$/i.test(String(r[0] ?? '').trim())) { brokerHeaderIdx = i; break }
  }
  const brokerHeadersRaw = ((summaryRows[brokerHeaderIdx] as unknown[]) || []).map(h => String(h ?? '').trim())
  while (brokerHeadersRaw.length > 0 && !brokerHeadersRaw[brokerHeadersRaw.length - 1]) brokerHeadersRaw.pop()
  const numCols = brokerHeadersRaw.length || 6
  let totalColIdx = brokerHeadersRaw.findIndex(h => /^total$/i.test(h))
  if (totalColIdx < 0) totalColIdx = numCols - 1

  const allBrokerRows: unknown[][] = []
  for (let i = brokerHeaderIdx + 1; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && row[0] && String(row[0]).trim()) allBrokerRows.push(row)
  }
  const dataRows = allBrokerRows.filter(r => String((r as unknown[])[0]).toLowerCase() !== 'totals')
  const totalsRow = allBrokerRows.find(r => String((r as unknown[])[0]).toLowerCase() === 'totals')

  const projectLabel = lease ? `Lease No. ${lease}` : ''
  const brokerHeaders = [...brokerHeadersRaw, 'Project']
  const totalColsWide = Math.max(brokerHeaders.length, 8)

  // ---- Borders ----
  const thin = { style: 'thin' as const, color: { argb: XLSX_COLOR_BLACK } }
  const borderAll = { top: thin, right: thin, bottom: thin, left: thin }

  // ---- Workbook ----
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Summary', {
    views: [{ showGridLines: false }],
    pageSetup: {
      orientation: 'landscape',
      paperSize: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    },
  })

  // ---- Column widths ----
  for (let i = 0; i < totalColsWide; i++) {
    let width = 11
    if (i === 0) width = 13
    else if (i === 1) width = 9
    else if (i === 2) width = 12
    else if (i === totalColIdx) width = 12
    else if (i === brokerHeaders.length - 1) width = 20
    ws.getColumn(i + 1).width = width
  }

  const centerAll = { horizontal: 'center' as const, vertical: 'middle' as const }

  // ---- BOP header block (rows 1-4) ----
  const bopLines = [
    { text: BOP.name, size: 14, bold: true },
    { text: BOP.address, size: 10, bold: false },
    { text: BOP.city, size: 10, bold: false },
    { text: BOP.phone, size: 10, bold: false },
  ]
  bopLines.forEach((line, i) => {
    ws.mergeCells(i + 1, 1, i + 1, totalColsWide)
    const cell = ws.getCell(i + 1, 1)
    cell.value = line.text
    cell.font = { name: 'Helvetica', size: line.size, bold: line.bold }
    cell.alignment = centerAll
  })

  // ---- INVOICE title (row 6) ----
  ws.mergeCells(6, 1, 6, totalColsWide)
  const invCell = ws.getCell(6, 1)
  invCell.value = 'INVOICE'
  invCell.font = { name: 'Helvetica', size: 16, bold: true }
  invCell.alignment = centerAll

  // ---- Horizontal rule (row 7 bottom border) ----
  for (let c = 1; c <= totalColsWide; c++) {
    ws.getCell(7, c).border = { bottom: { style: 'medium', color: { argb: XLSX_COLOR_BLACK } } }
  }

  const labelFont = { name: 'Helvetica', size: 10, bold: true, color: { argb: XLSX_COLOR_GRAY } }
  const valueFont = { name: 'Helvetica', size: 10, color: { argb: XLSX_COLOR_BLACK } }
  const boldValueFont = { name: 'Helvetica', size: 10, bold: true, color: { argb: XLSX_COLOR_BLACK } }

  const leftLabelCol = 1
  const leftValueColStart = 2
  const rightLabelCol = Math.floor(totalColsWide / 2) + 1
  const rightValueColStart = rightLabelCol + 1

  const setLabelLeft = (row: number, text: string) => {
    const c = ws.getCell(row, leftLabelCol)
    c.value = text
    c.font = labelFont
    c.alignment = { horizontal: 'left', vertical: 'top' }
  }
  const setValueLeft = (row: number, text: string, bold = false) => {
    const startCol = leftValueColStart
    const endCol = rightLabelCol - 1
    if (endCol > startCol) ws.mergeCells(row, startCol, row, endCol)
    const c = ws.getCell(row, startCol)
    c.value = text
    c.font = bold ? boldValueFont : valueFont
    c.alignment = { horizontal: 'left', vertical: 'top', wrapText: true }
  }
  const setLabelRight = (row: number, text: string) => {
    const c = ws.getCell(row, rightLabelCol)
    c.value = text
    c.font = labelFont
    c.alignment = { horizontal: 'left', vertical: 'top' }
  }
  const setValueRight = (row: number, text: string) => {
    const startCol = rightValueColStart
    const endCol = totalColsWide
    if (endCol > startCol) ws.mergeCells(row, startCol, row, endCol)
    const c = ws.getCell(row, startCol)
    c.value = text
    c.font = valueFont
    c.alignment = { horizontal: 'left', vertical: 'top', wrapText: true }
  }

  // Row 9: Date + Invoice #
  setLabelLeft(9, 'Date:')
  setValueLeft(9, invoiceDate)
  setLabelRight(9, 'Invoice #:')
  setValueRight(9, invoiceNum)

  // Rows 11-15: Bill To on left, Lease/PID/County/Unit/Type on right
  setLabelLeft(11, 'Bill To:')
  setValueLeft(11, BILL_TO.company, true)
  setValueLeft(12, BILL_TO.attn)
  setValueLeft(13, BILL_TO.address)
  setValueLeft(14, BILL_TO.city)

  setLabelRight(11, 'Lease No.:')
  setValueRight(11, lease)
  setLabelRight(12, 'PID:')
  setValueRight(12, parcel)
  setLabelRight(13, 'County:')
  setValueRight(13, county)
  setLabelRight(14, 'Unit:')
  setValueRight(14, unit)
  setLabelRight(15, 'Type:')
  setValueRight(15, workType)

  // Row 17: Period + DUE UPON RECEIPT
  setLabelLeft(17, 'Period:')
  setValueLeft(17, period)
  const dueCell = ws.getCell(17, totalColsWide)
  dueCell.value = 'DUE UPON RECEIPT'
  dueCell.font = { name: 'Helvetica', size: 10, bold: true, color: { argb: XLSX_COLOR_RED } }
  dueCell.alignment = { horizontal: 'right', vertical: 'middle' }

  // Row 19: Broker table header
  const brokerHeaderRow = 19
  brokerHeaders.forEach((h, i) => {
    const cell = ws.getCell(brokerHeaderRow, i + 1)
    cell.value = h
    cell.font = { name: 'Helvetica', size: 9, bold: true, color: { argb: XLSX_COLOR_BLACK } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XLSX_COLOR_PINK } }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = borderAll
  })
  ws.getRow(brokerHeaderRow).height = 36

  // Broker data + totals rows
  let rowIdx = brokerHeaderRow
  const writeBrokerRow = (row: unknown[], isTotalsRow: boolean) => {
    rowIdx++
    for (let ci = 0; ci < brokerHeaders.length; ci++) {
      const cell = ws.getCell(rowIdx, ci + 1)
      const h = brokerHeaders[ci]
      const isProjectCol = ci === brokerHeaders.length - 1
      const isTotalCol = ci === totalColIdx
      let value: unknown = ''
      if (isProjectCol) value = isTotalsRow ? '' : projectLabel
      else if (ci < numCols) {
        const v = row[ci]
        value = v === undefined || v === null ? '' : v
      }
      if (isTotalsRow && ci === 0) value = 'Totals'

      cell.value = value as ExcelJS.CellValue
      const kind = classifyHeader(h, ci === 0)
      cell.font = { name: 'Helvetica', size: 9, bold: isTotalsRow, color: { argb: XLSX_COLOR_BLACK } }
      cell.border = borderAll
      cell.alignment = {
        horizontal: ci === 0 ? 'left' : (kind === 'text' ? 'left' : (kind === 'days' || kind === 'miles' ? 'center' : 'right')),
        vertical: 'middle',
        wrapText: isProjectCol,
      }
      if (typeof value === 'number') {
        if (kind === 'currency') cell.numFmt = XLSX_CURRENCY_FMT
        else if (kind === 'days') cell.numFmt = XLSX_DAYS_FMT
        else if (kind === 'miles') cell.numFmt = '0'
      }
      if (isTotalsRow && isTotalCol) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XLSX_COLOR_YELLOW } }
      }
    }
  }

  for (const dr of dataRows) writeBrokerRow(dr as unknown[], false)
  if (totalsRow) writeBrokerRow(totalsRow as unknown[], true)

  // Footer note
  const noteRow = rowIdx + 2
  ws.mergeCells(noteRow, 1, noteRow, totalColsWide)
  const nc = ws.getCell(noteRow, 1)
  nc.value = 'Please contact our accounting department with any questions regarding invoices'
  nc.font = { name: 'Helvetica', size: 8, italic: true, color: { argb: XLSX_COLOR_GRAY } }
  nc.alignment = { horizontal: 'center' }

  // ---- Work Detail sheet ----
  if (detailSheet) {
    const detRows = XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][]
    let detHeaderIdx = 1
    for (let i = 0; i < Math.min(detRows.length, 5); i++) {
      const r = detRows[i] as unknown[]
      if (r && /^landman$/i.test(String(r[0] ?? '').trim())) { detHeaderIdx = i; break }
    }
    const detailHeaders = ((detRows[detHeaderIdx] as unknown[]) || []).map(h => String(h ?? '').trim())
    while (detailHeaders.length > 0 && !detailHeaders[detailHeaders.length - 1]) detailHeaders.pop()
    const detailData = detRows.slice(detHeaderIdx + 1).filter(r => {
      const row = r as unknown[]
      return row && row[0] && String(row[0]).trim() !== ''
    })

    if (detailHeaders.length > 0) {
      const dws = wb.addWorksheet('Work Detail', {
        views: [{ showGridLines: false }],
        pageSetup: {
          orientation: 'landscape', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
          margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
        },
      })

      detailHeaders.forEach((h, i) => {
        const lo = h.toLowerCase()
        let w = 12
        if (/description|complete/.test(lo) && !/miles\s+description/.test(lo)) w = 40
        else if (/miles\s+description/.test(lo)) w = 18
        else if (/landman|prospect/.test(lo)) w = 14
        else if (/^date$/.test(lo)) w = 11
        else if (/legal/.test(lo)) w = 18
        else if (/focus|lease/.test(lo)) w = 13
        else if (/\bdays\b/.test(lo)) w = 7
        else if (/\bmiles\b/.test(lo) && !/mileage|description/.test(lo)) w = 7
        dws.getColumn(i + 1).width = w
      })

      // "Work Detail" title (row 1)
      dws.mergeCells(1, 1, 1, detailHeaders.length)
      const t = dws.getCell(1, 1)
      t.value = 'Work Detail'
      t.font = { name: 'Helvetica', size: 12, bold: true }
      t.alignment = { horizontal: 'left', vertical: 'middle' }
      dws.getRow(1).height = 22
      for (let c = 1; c <= detailHeaders.length; c++) {
        dws.getCell(1, c).border = { bottom: { style: 'medium', color: { argb: XLSX_COLOR_BLACK } } }
      }

      // Header row (row 3)
      const dHeaderRow = 3
      detailHeaders.forEach((h, i) => {
        const cell = dws.getCell(dHeaderRow, i + 1)
        cell.value = h
        cell.font = { name: 'Helvetica', size: 8, bold: true }
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XLSX_COLOR_PINK } }
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        cell.border = borderAll
      })
      dws.getRow(dHeaderRow).height = 32

      const detailTotalColIdx = detailHeaders.findIndex(h => /^total$/i.test(h))
      const isRateColumn = (h: string) =>
        /\brate\b|dayrate|per\s+(day|diem|hour|hr|mile)|amt\.?\s*per|amount\s*per/i.test(h)

      const totalsRowDetail: unknown[] = new Array(detailHeaders.length).fill('')
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
          if (any) totalsRowDetail[i] = sum
        }
      })

      let dRow = dHeaderRow
      const writeDetailRow = (row: unknown[], isTotalsRow: boolean) => {
        dRow++
        for (let ci = 0; ci < detailHeaders.length; ci++) {
          const cell = dws.getCell(dRow, ci + 1)
          const h = detailHeaders[ci]
          const v = row[ci]
          const value: unknown = v === undefined || v === null ? '' : v
          cell.value = value as ExcelJS.CellValue
          const kind = classifyHeader(h, ci === 0)
          cell.font = { name: 'Helvetica', size: 8, bold: isTotalsRow }
          cell.border = borderAll
          cell.alignment = {
            horizontal: kind === 'currency' ? 'right' : (kind === 'days' || kind === 'miles' ? 'center' : 'left'),
            vertical: 'top', wrapText: true,
          }
          if (typeof value === 'number') {
            if (kind === 'currency') cell.numFmt = XLSX_CURRENCY_FMT
            else if (kind === 'days') cell.numFmt = XLSX_DAYS_FMT
            else if (kind === 'miles') cell.numFmt = '0'
          }
          if (isTotalsRow && ci === detailTotalColIdx) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XLSX_COLOR_YELLOW } }
          }
        }
      }

      for (const dr of detailData) writeDetailRow(dr as unknown[], false)
      writeDetailRow(totalsRowDetail, true)
    }
  }

  const arrayBuf = await wb.xlsx.writeBuffer()
  return Buffer.from(arrayBuf as ArrayBuffer)
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
          // Build a fresh formatted Excel invoice mirroring the PDF layout.
          const xlsxBuffer = await buildInvoiceXlsx(excelBuffer, billing, invoiceDateOverride)
          zip.file(`${outputName}.xlsx`, xlsxBuffer)
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
