import { signOut } from 'firebase/auth'
import { auth } from '../firebase'
import { airports } from '../data/airports'

export default function Dashboard({ airline }) {
  const home = airports.find((a) => a.code === airline.homeAirport)

  return (
    <div className="dashboard">
      <header>
        <h1>{airline.name}</h1>
        <button type="button" className="link-btn" onClick={() => signOut(auth)}>
          Odhlásit se
        </button>
      </header>

      <div className="stats">
        <div className="stat">
          <span className="label">Hotovost</span>
          <span className="value">{airline.cash.toLocaleString('cs-CZ')} Kč</span>
        </div>
        <div className="stat">
          <span className="label">Domovské letiště</span>
          <span className="value">
            {airline.homeAirport}
            {home ? ` – ${home.city}` : ''}
          </span>
        </div>
      </div>

      <p className="placeholder-note">
        Další moduly (flotila, trasy, cestující, finance) přibudou postupně.
      </p>
    </div>
  )
}
