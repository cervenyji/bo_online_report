# Airline Simulator (online, multiplayer)

Simulátor provozování letecké společnosti hraný online s kamarády. Vyvíjí se
postupně po modulech. Frontend: React + Vite. Backend/data: Firebase
(Authentication + Firestore), oboje zdarma v rámci Spark planu.

## Aktuální stav

**Modul 1: Účty + založení aerolinky**

- Registrace/přihlášení (e-mail+heslo nebo Google)
- Založení vlastní letecké společnosti (název, domovské letiště, startovní kapitál)
- Základní dashboard s přehledem aerolinky

Data se ukládají do Firestore kolekce `airlines`, jeden dokument na hráče
(ID dokumentu = UID hráče).

## Založení vlastního Firebase projektu (zdarma)

1. Jdi na [console.firebase.google.com](https://console.firebase.google.com) a vytvoř nový projekt.
2. V sekci **Build → Authentication** povol metody přihlášení: **Email/Password** a volitelně **Google**.
3. V sekci **Build → Firestore Database** vytvoř databázi (produkční režim).
4. Nasaď pravidla ze souboru `firestore.rules` (Firestore → Rules → vlož obsah souboru → Publish).
5. V **Project settings → General → Your apps** přidej webovou aplikaci a zkopíruj konfigurační hodnoty.
6. Zkopíruj `.env.example` do `.env` a vyplň hodnoty z kroku 5:

   ```
   cp .env.example .env
   ```

## Lokální vývoj

```bash
npm install
npm run dev
```

## Nasazení zdarma (Firebase Hosting)

```bash
npm run build
npm install -g firebase-tools
firebase login
firebase init hosting   # public directory: dist, single-page app: yes
firebase deploy
```

Kamarádi pak hru otevřou na URL, kterou `firebase deploy` vypíše
(`https://<projekt>.web.app`), a každý si vytvoří vlastní účet a aerolinku.

## Plán dalších modulů

- Flotila letadel (nákup/pronájem, parametry)
- Letecké trasy mezi letišti
- Cestující, poptávka, ceny letenek
- Finance, náklady, výnosy, herní čas (tiky)
- Žebříček hráčů / soutěžní prvky
