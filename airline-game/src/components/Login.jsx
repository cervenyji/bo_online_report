import { useState } from 'react'
import {
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  signInWithPopup,
} from 'firebase/auth'
import { auth, googleProvider } from '../firebase'

export default function Login() {
  const [mode, setMode] = useState('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  async function handleSubmit(e) {
    e.preventDefault()
    setError('')
    setBusy(true)
    try {
      if (mode === 'login') {
        await signInWithEmailAndPassword(auth, email, password)
      } else {
        await createUserWithEmailAndPassword(auth, email, password)
      }
    } catch (err) {
      setError(err.message)
    } finally {
      setBusy(false)
    }
  }

  async function handleGoogleSignIn() {
    setError('')
    setBusy(true)
    try {
      await signInWithPopup(auth, googleProvider)
    } catch (err) {
      setError(err.message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="auth-card">
      <h1>Airline Simulator</h1>
      <p className="subtitle">Přihlas se a začni budovat svoji leteckou společnost</p>

      <form onSubmit={handleSubmit}>
        <label>
          E-mail
          <input
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
          />
        </label>
        <label>
          Heslo
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            minLength={6}
            required
          />
        </label>
        {error && <p className="error">{error}</p>}
        <button type="submit" disabled={busy}>
          {mode === 'login' ? 'Přihlásit se' : 'Registrovat se'}
        </button>
      </form>

      <button
        type="button"
        className="google-btn"
        onClick={handleGoogleSignIn}
        disabled={busy}
      >
        Pokračovat přes Google
      </button>

      <button
        type="button"
        className="link-btn"
        onClick={() => setMode(mode === 'login' ? 'signup' : 'login')}
      >
        {mode === 'login' ? 'Nemáš účet? Zaregistruj se' : 'Už máš účet? Přihlas se'}
      </button>
    </div>
  )
}
