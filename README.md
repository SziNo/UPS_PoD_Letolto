# UPS PoD Letöltő

Automatizált PowerShell script UPS Proof of Delivery (PoD) dokumentumok letöltéséhez.

## 📋 Tartalomjegyzék
- [Funkciók](#funkciók)
- [Rendszerkövetelmények](#rendszerkövetelmények)
- [Telepítés lépésről lépésre](#telepítés-lépésről-lépésre)
  - [1. Python telepítése](#1-python-telepítése)
  - [2. Python csomagok telepítése](#2-python-csomagok-telepítése)
  - [3. Script letöltése GitHub-ról](#3-script-letöltése-github-ról)
  - [4. Fájl feloldása (unblock)](#4-fájl-feloldása-unblock)
- [Program futtatása](#program-futtatása)
- [Első használat lépésről lépésre](#első-használat-lépésről-lépésre)
- [Eredmények](#eredmények)
- [Megjegyzések](#megjegyzések)
- [Hibaelhárítás](#hibaelhárítás)
- [Kapcsolódó fájlok](#kapcsolódó-fájlok)

---

## ✨ Funkciók

- ✅ Automatikus bejelentkezés UPS fiókba (egyszer kell megadni)
- ✅ Tracking számok beolvasása Excel fájlból
- ✅ PoD dokumentumok letöltése PDF formátumban
- ✅ Fájlok automatikus átnevezése az „összefűz" oszlop alapján
- ✅ Feldolgozott sorok zöld színezése Excel-ben (A–E oszlopok)
- ✅ Részletes naplózás a GUI-ban
- ✅ Futás közben megállítható (STOP gomb)

---

## 💻 Rendszerkövetelmények

- **Windows 10 vagy 11**
- **Google Chrome** böngésző (telepítve)
- **Python 3.11 vagy újabb**
- **PowerShell** (alapból része a Windows-nak)
- **Internet kapcsolat** (VPN használata esetén lehet, hogy ki kell kapcsolni)

---

## 🔧 Telepítés lépésről lépésre

### 1. Python telepítése

1. Nyisd meg: [https://www.python.org/downloads](https://www.python.org/downloads)
2. **Fontos!** Van egy nagy sárga „Download Python install manager" gomb – **NE ARRA KATTINTS!**
3. Kattints alatta az **„Or get the standalone installer for Python 3.x.x"** linkre
4. A telepítő futtatásakor az első képernyőn **PIPÁLD BE**: `Add Python.exe to PATH`
   > Ez nagyon fontos, mert ettől fog működni a `python` és `pip` parancs PowerShellben
5. Kattints: **Install Now**

### 2. Python csomagok telepítése

Nyiss egy PowerShell ablakot, és futtasd:

```powershell
pip install pandas openpyxl selenium
```

> **Megjegyzés:** Ha VPN-t használsz, előtte kapcsold ki, mert a VPN blokkolhatja a csomagok letöltését.

### 3. Script letöltése GitHub-ról

1. Nyisd meg a repository GitHub oldalát
2. Kattints a zöld **„Code"** gombra
3. Válaszd a **„Download ZIP"** lehetőséget
4. Csomagold ki a letöltött mappát `UPS_PoD_Letolto-main`
5. Helyezd el a `C:\temp` mappába 

### 4. Fájl feloldása (unblock)

Mielőtt futtatnád, fel kell oldanod a letöltött PowerShell fájlt:

1. Jobb klikk a `UPS_PoD_Letolto.ps1` fájlon → **Properties / Tulajdonságok**
2. Ha van **„Unblock"** vagy **„Tiltás feloldása"** checkbox, **PIPÁLD BE!**
3. Kattints **Alkalmaz**, majd **OK**

---

## 🚀 Program futtatása

Nyiss egy PowerShell ablakot, majd futtasd az alábbi parancsot:

```powershell
powershell -ExecutionPolicy Bypass -File C:\temp\UPS_PoD_Letolto-main\UPS_PoD_Letolto.ps1
```

Sikeres indítás után megjelenik a program grafikus felülete.

---

## 📝 Első használat lépésről lépésre

### 1️⃣ PoD Chrome indítása

1. Kattints a **„PoD Chrome indítása"** gombra a GUI-ban
2. Megnyílik egy új Chrome ablak
3. Jelentkezz be az UPS fiókodba ebben az ablakban

> ⚠️ **FONTOS: EZT AZ ABLAKOT NE ZÁRD BE!** A script ehhez fog csatlakozni.

### 2️⃣ Adatok megadása

| Mező | Leírás |
|------|--------|
| **UPS URL** | Alapból meg van adva. Csak akkor módosítsd, ha az UPS megváltoztatja a tracking oldal címét. |
| **Excel fájl** | Tallózással válaszd ki a tracking számokat tartalmazó Excel fájlt |
| **Letöltési mappa** | Válaszd ki, hová szeretnéd menteni a letöltött PDF fájlokat |

### 3️⃣ Excel fájl előkészítése

A fájlnak tartalmaznia kell az alábbi oszlopokat:

| Oszlop | Leírás |
|--------|--------|
| `Tracking Number` | A nyomkövetési számok |
| `összefűz` | A letöltött fájlok végső neve **(ű-VEL!)** |

### 4️⃣ Letöltés indítása

1. Ha mindent beállítottál, kattints az **„Indítás"** gombra
2. A script elkezdi feldolgozni a sorokat
3. A naplóban követheted a folyamatot
4. Ha szükséges, a **„STOP"** gombbal bármikor megszakíthatod

---

## 📂 Eredmények

### Letöltött PDF fájlok
- A PDF-ek a megadott **letöltési mappába** kerülnek
- A fájlnevek az `összefűz` oszlop értékei lesznek

### Feldolgozott Excel
- A script létrehoz egy új Excel fájlt: `[eredeti_név]_FELDOLGOZOTT.xlsx`
- Ez szintén a **letöltési mappába** kerül
- A sikeresen letöltött sorokban az **A–E oszlopok zöld színt kapnak**

---

## 💡 Megjegyzések

- **VPN használata:** Ha VPN-nel dolgozol, előfordulhat, hogy a script nem működik. Ilyenkor kapcsold ki a VPN-t a futtatás idejére.
- **Csak egyszer kell bejelentkezni:** A „PoD Chrome" eltárolja a bejelentkezési adatokat, így legközelebb már nem kell újra megadnod.
- **POD nem elérhető:** Ha egy tracking számhoz még nincs feltöltve a PoD az UPS rendszerébe, a script kihagyja azt a sort (nem színezi be). Következő futtatáskor újra megpróbálja.

---

## ❓ Hibaelhárítás

| Hiba | Megoldás |
|------|----------|
| **A program nem indul** | Jobb klikk → Properties → „Unblock" checkbox bepipálása |
| **Python vagy pip nem található** | Ellenőrizd, hogy telepítéskor be lett-e pipálva az „Add Python to PATH" |
| **Csomagok nem települnek** | Kapcsold ki a VPN-t, majd: `pip install pandas openpyxl selenium` |
| **„A POD Chrome nem fut" hiba** | Először mindig a „PoD Chrome indítása" gombra kell kattintani és bejelentkezni |
| **Script elakad** | Ellenőrizd, hogy a PoD Chrome ablak nyitva van-e és be vagy-e jelentkezve |

---

## 📦 Kapcsolódó fájlok

- `UPS_PoD_Letolto.ps1` – a fő PowerShell script
- `README.md` – ez a dokumentáció

> A Python script ideiglenes fájlként jön létre futáskor, nem kell külön telepíteni.

---

*Verzió: 1.0 | Utolsó frissítés: 2026. március*
