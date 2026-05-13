import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'BOP Range Resources Invoice Generator',
  description: 'BOP Range Resources Title Invoice Generator',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body style={{ margin: 0, padding: 0 }}>{children}</body>
    </html>
  )
}
