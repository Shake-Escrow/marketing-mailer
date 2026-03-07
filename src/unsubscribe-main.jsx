import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import UnsubscribeApp from './UnsubscribeApp'
import './index.css'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <UnsubscribeApp />
  </StrictMode>,
)