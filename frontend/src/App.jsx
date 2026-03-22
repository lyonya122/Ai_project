import { useState } from 'react'

const API_URL = 'http://127.0.0.1:8000'

export default function App() {
  const [topic, setTopic] = useState('AI-приложение для генерации презентаций')
  const [audience, setAudience] = useState('Преподаватель и комиссия')
  const [purpose, setPurpose] = useState('Защитить проект и показать прикладную ценность')
  const [slideCount, setSlideCount] = useState(8)
  const [status, setStatus] = useState('')
  const [loading, setLoading] = useState(false)

  const uploadKnowledge = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setStatus('Загружаю базу знаний...')
    const formData = new FormData()
    formData.append('file', file)
    const response = await fetch(`${API_URL}/api/knowledge/upload`, {
      method: 'POST',
      body: formData,
    })
    const data = await response.json()
    if (!response.ok) {
      setStatus(data.detail || 'Ошибка загрузки базы знаний')
      return
    }
    setStatus(`База знаний загружена: ${data.chunks_added} чанков`)
  }

  const generatePresentation = async () => {
    setLoading(true)
    setStatus('Генерирую презентацию...')
    try {
      const response = await fetch(`${API_URL}/api/presentations/generate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          topic,
          audience,
          purpose,
          slide_count: Number(slideCount),
          tone: 'professional',
        }),
      })

      if (!response.ok) {
        const err = await response.json().catch(() => ({}))
        throw new Error(err.detail || 'Не удалось сгенерировать презентацию')
      }

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'generated_presentation.pptx'
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
      setStatus('Готово: презентация скачана.')
    } catch (error) {
      setStatus(error.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="page">
      <div className="card">
        <h1>AI Presentation Generator</h1>
        <p className="sub">Три агента: planner → writer → designer + RAG + экспорт в PPTX</p>

        <label>Тема</label>
        <textarea value={topic} onChange={(e) => setTopic(e.target.value)} rows={4} />

        <label>Целевая аудитория</label>
        <input value={audience} onChange={(e) => setAudience(e.target.value)} />

        <label>Цель презентации</label>
        <input value={purpose} onChange={(e) => setPurpose(e.target.value)} />

        <label>Количество слайдов</label>
        <input type="number" min="4" max="15" value={slideCount} onChange={(e) => setSlideCount(e.target.value)} />

        <label>Загрузить базу знаний (pdf/txt)</label>
        <input type="file" onChange={uploadKnowledge} />

        <div className="actions">
          <button disabled={loading} onClick={generatePresentation}>
            {loading ? 'Генерация...' : 'Сгенерировать'}
          </button>
        </div>

        {status && <div className="status">{status}</div>}
      </div>
    </div>
  )
}
