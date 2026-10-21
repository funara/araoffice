import { existsSync, writeFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import sharp from 'sharp'

const __dirname = dirname(fileURLToPath(import.meta.url))
const root = resolve(__dirname, '..')

const sourcePng = resolve(root, 'public/araoffice-icon.png')
const targetShellIco = resolve(root, 'apps/shell/build/icon.ico')
const targetShellPng = resolve(root, 'apps/shell/build/icon.png')

if (!existsSync(sourcePng)) {
  console.error(`Source icon not found at ${sourcePng}`)
  process.exit(1)
}

async function generateAppIcons() {
  console.log(`Generating app icons from ${sourcePng}...`)

  // 2. Generate multi-resolution icon sizes for ICO
  const sizes = [16, 24, 32, 48, 64, 128, 256]
  const pngBuffers = []

  for (const size of sizes) {
    const buf = await sharp(sourcePng)
      .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png({ compressionLevel: 9 })
      .toBuffer()
    pngBuffers.push({ size, buf })
  }

  // 3. Write 256x256 PNG to apps/shell/build/icon.png
  const png256 = pngBuffers.find((p) => p.size === 256).buf
  writeFileSync(targetShellPng, png256)
  console.log(`Updated ${targetShellPng}`)

  // 4. Construct Windows ICO container
  const headerLen = 6
  const dirEntryLen = 16
  const totalDirLen = headerLen + sizes.length * dirEntryLen

  let currentOffset = totalDirLen
  const dirEntries = []

  for (const item of pngBuffers) {
    const b = item.buf
    const entry = Buffer.alloc(dirEntryLen)
    entry.writeUInt8(item.size === 256 ? 0 : item.size, 0) // width
    entry.writeUInt8(item.size === 256 ? 0 : item.size, 1) // height
    entry.writeUInt8(0, 2) // color count
    entry.writeUInt8(0, 3) // reserved
    entry.writeUInt16LE(1, 4) // color planes
    entry.writeUInt16LE(32, 6) // bits per pixel
    entry.writeUInt32LE(b.length, 8) // size of image data
    entry.writeUInt32LE(currentOffset, 12) // offset
    dirEntries.push(entry)
    currentOffset += b.length
  }

  const header = Buffer.alloc(headerLen)
  header.writeUInt16LE(0, 0) // reserved
  header.writeUInt16LE(1, 2) // type: 1 for icon
  header.writeUInt16LE(sizes.length, 4) // image count

  const finalIco = Buffer.concat([header, ...dirEntries, ...pngBuffers.map((p) => p.buf)])
  writeFileSync(targetShellIco, finalIco)
  console.log(`Updated ${targetShellIco} (${finalIco.length} bytes, ${sizes.length} frames)`)
}

generateAppIcons().catch((err) => {
  console.error('Failed to generate app icons:', err)
  process.exit(1)
})
