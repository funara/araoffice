import type { Lang } from '@araoffice/i18n'
import type { DiagramShapeType } from './diagramCommands'

type ShapeHelp = { label: string; description: string }

const english: Record<DiagramShapeType, ShapeHelp> = {
  rectangle: { label: 'Rectangle', description: 'A general-purpose step or container.' },
  rounded: { label: 'Process', description: 'An action or process step.' },
  ellipse: { label: 'Start / End', description: 'The start or end of a flow.' },
  doubleEllipse: { label: 'Double ellipse', description: 'A highlighted endpoint.' },
  diamond: { label: 'Decision', description: 'A condition that branches the flow.' },
  cylinder: { label: 'Database', description: 'Persistent data storage.' },
  text: { label: 'Text', description: 'A free-form annotation.' },
  triangle: { label: 'Triangle', description: 'A directional or warning point.' },
  hexagon: { label: 'Preparation', description: 'A preparation or setup step.' },
  cloud: { label: 'Cloud', description: 'An external service or network.' },
}

const indonesian: Record<DiagramShapeType, ShapeHelp> = {
  rectangle: { label: 'Persegi panjang', description: 'Langkah atau wadah serbaguna.' },
  rounded: { label: 'Proses', description: 'Langkah aksi atau proses.' },
  ellipse: { label: 'Mulai / selesai', description: 'Awal atau akhir alur.' },
  doubleEllipse: { label: 'Elips ganda', description: 'Titik akhir yang ditonjolkan.' },
  diamond: { label: 'Keputusan', description: 'Kondisi yang mencabangkan alur.' },
  cylinder: { label: 'Database', description: 'Penyimpanan data permanen.' },
  text: { label: 'Teks', description: 'Anotasi bebas.' },
  triangle: { label: 'Segitiga', description: 'Penanda arah atau peringatan.' },
  hexagon: { label: 'Persiapan', description: 'Langkah persiapan atau penyiapan.' },
  cloud: { label: 'Cloud', description: 'Layanan atau jaringan eksternal.' },
}

export function shapeHelp(type: DiagramShapeType, lang: Lang): ShapeHelp {
  return (lang === 'id' ? indonesian : english)[type]
}

export function quickAddCopy(lang: Lang) {
  if (lang === 'id') {
    return {
      title: 'Tambah node terhubung',
      search: 'Cari bentuk',
      noMatch: 'Bentuk tidak ditemukan',
      hint: 'Enter untuk tambah · Esc atau klik di luar untuk batal',
    }
  }
  return {
    title: 'Add connected node',
    search: 'Search shapes',
    noMatch: 'No matching shape',
    hint: 'Enter to add · Esc or click outside to cancel',
  }
}
