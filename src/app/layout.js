export const metadata = {
  title: 'KPI Agent – FiberNC',
  description: 'Techniker KPI Kontrolle',
}

export default function RootLayout({ children }) {
  return (
    <html lang="de">
      <body style={{ margin: 0, padding: 0, background: '#0a0e1a' }}>{children}</body>
    </html>
  )
}
