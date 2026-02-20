// src/components/Header.jsx
import { loginRequest } from '../authConfig'
import styles from './Header.module.css' // or however you import styles

export default function Header({ account, isAuthenticated, instance }) {
  const handleSignOut = () => {
    instance.logoutPopup({ postLogoutRedirectUri: window.location.origin })
  }

  return (
    <header className={styles.header}>
      <div className={styles.logo}>Marketing Mailer</div>
      {isAuthenticated && account && (
        <div className={styles.userArea}>
          <span className={styles.username}>{account.username}</span>
          <button className={styles.signOutBtn} onClick={handleSignOut}>
            Sign Out
          </button>
        </div>
      )}
    </header>
  )
}