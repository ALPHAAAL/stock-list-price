import './globals.css';

export const metadata = {
  title: 'Stock app',
  description: 'Created by NOT YOU',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
