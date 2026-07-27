import { useState } from 'react'
import { doc, setDoc, serverTimestamp } from 'firebase/firestore'
import { db } from '../firebase'
import { useAuth } from '../context/AuthContext'
import { airports } from '../data/airports'

const STARTING_CASH = 10_000_000

export default function CreateAirline() {
  const { user } = useAuth()
  const [name, setName] = useState('')
  const [homeAirport, setHomeAirport] = useState(airports[0].code)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  async function handleSubmit(e) {
    e.preventDefault()
    setError('')
    if (name.trim().length < 3) {
      setError('Název aerolinky musí mít alespoň 3 znaky.')
      return
    }
    setBusy(true)
    try {
      await setDoc(doc(db, 'airlines', user.uid), {
        name: name.trim(),
        homeAirport,
        cash: STARTING_CASH,
        ownerUid: user.uid,
        ownerEmail: user.email,
        createdAt: serverTimestamp(),
      })
    } catch (err) {
      setError(err.message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="auth-card">
      <h1>Založ svoji aerolinku</h1>
      <p className="subtitle">
        Startovní kapitál: {STARTING_CASH.toLocaleString('cs-CZ')} Kč
      </p>

      <form onSubmit={handleSubmit}>
        <label>
          Název aerolinky
          <input
            type="text"
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="např. Czech Wings"
            required
          />
        </label>
        <label>
          Domovské letiště
          <select
            value={homeAirport}
            onChange={(e) => setHomeAirport(e.target.value)}
          >
            {airports.map((a) => (
              <option key={a.code} value={a.code}>
                {a.code} – {a.city} ({a.country})
              </option>
            ))}
          </select>
        </label>
        {error && <p className="error">{error}</p>}
        <button type="submit" disabled={busy}>
          Založit aerolinku
        </button>
      </form>
    </div>
  )
}
