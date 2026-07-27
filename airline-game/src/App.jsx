import { useEffect, useState } from 'react'
import { doc, onSnapshot } from 'firebase/firestore'
import { db } from './firebase'
import { AuthProvider, useAuth } from './context/AuthContext'
import Login from './components/Login'
import CreateAirline from './components/CreateAirline'
import Dashboard from './components/Dashboard'
import './App.css'

function AppContent() {
  const { user, loading } = useAuth()
  const [airline, setAirline] = useState(undefined)

  useEffect(() => {
    if (!user) {
      setAirline(undefined)
      return
    }
    return onSnapshot(doc(db, 'airlines', user.uid), (snap) => {
      setAirline(snap.exists() ? snap.data() : null)
    })
  }, [user])

  if (loading) return <div className="center-screen">Načítám…</div>
  if (!user) return <Login />
  if (airline === undefined) return <div className="center-screen">Načítám aerolinku…</div>
  if (airline === null) return <CreateAirline />
  return <Dashboard airline={airline} />
}

export default function App() {
  return (
    <AuthProvider>
      <AppContent />
    </AuthProvider>
  )
}
