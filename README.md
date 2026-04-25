# eFactura RO — Generator XML (Excel + VBA)

Generator gratuit, open-source, de fișiere XML conforme **CIUS-RO 1.0.1** pentru **RO e-Factura (ANAF)**, pornind de la un workbook Excel.

> **Versiune:** 3.0.0 &nbsp; **Licență:** [MIT](#licență)

---

## Pentru cine este

Persoane care intră obligatoriu în RO e-Factura din **1 iunie 2026**:

- drepturi de autor
- profesori cu meditații
- agricultori, persoane fizice cu activitate independentă identificate prin CNP
- PFA, II, microîntreprinderi mici (plătitoare sau neplătitoare TVA)

Un singur workbook poate deservi mai mulți emitenți.

## Ce face

- Generează fișiere XML UBL Invoice 2.1 / CIUS-RO 1.0.1 valide.
- Auto-detectează cazurile **B2B**, **B2C consumator** (cu anonimizare CNP la 13 zerouri conform GDPR), **PFA cu CNP plătitor TVA**.
- Furnizor plătitor sau neplătitor TVA — comutare prin coloana `PlatitorTVA` (DA/NU).
- **Storno**: factură tip 380 cu **cantitate negativă** pe linie (preț rămâne pozitiv).
- Reconciliere automată antet ↔ linii (toleranță 0.01) — facturile cu sume incoerente sunt raportate, nu generate.
- Fișiere `.xml` UTF-8 fără BOM (cum cere ANAF).

## Ce NU face

- **Nu trimite** fișierul în SPV — încărcarea în SPV-ul ANAF este responsabilitatea ta și necesită certificat digital calificat.
- Nu semnează electronic.
- Nu te înregistrează în Registrul RO e-Factura.
- v1: doar moneda RON, doar tip factură 380, cote TVA suportate 0/5/11/21.

---

## Instalare (în 5 pași)

1. Descarcă fișierele din acest repository: `eFactura_RO_v3.xlsx` și `generare_eFactura.bas`.
2. Deschide `eFactura_RO_v3.xlsx` în Excel.
3. **Fișier → Salvare ca → Registru de lucru Excel cu macrocomenzi (`*.xlsm`)**.
4. Apasă `Alt + F11` ca să deschizi editorul VBA.
5. **Fișier → Import File...** → selectează `generare_eFactura.bas` → Apasă `Open`. Salvează fișierul.

Gata. Ai un `.xlsm` complet funcțional.

> **Notă:** la prima rulare Excel poate cere să activezi conținutul / macrocomenzile. Trebuie să fie **„Activează tot conținutul"** (sau activarea macrourilor din Centrul de încredere).

---

## Cum se folosește (4 pași)

### 1. Foaia `Furnizori`

Completează datele firmei tale (sau ale fiecărui furnizor pe care îl deservești):

| Coloană | Ce pui |
|---|---|
| `RegistrationName` | Nume firmă / PFA, exact ca la ANAF |
| `CompanyID` | CIF/CUI — **fără** prefixul `RO` dacă nu ești plătitor; **cu** `RO` dacă ești plătitor |
| `PlatitorTVA` | `DA` sau `NU` — controlează blocul TVA în XML |
| `MotivScutire` | Doar dacă `PlatitorTVA = NU`. Lasă gol pentru valoarea implicită ("Entitatea nu este inregistrata in scopuri de TVA") |
| `IBAN`, `BankName` | Opționale — apar în blocul `PaymentMeans` din XML |

### 2. Foaia `Clienti`

Completează clienții. Macroul detectează automat tipul:

| Identificator client | Caz | Tratament |
|---|---|---|
| CIF cu `RO` (ex. `RO12345678`) | Persoană juridică plătitoare TVA | B2B normal |
| CIF fără `RO` (ex. `12345678`) | Persoană juridică neplătitoare TVA | B2B simplificat |
| CNP de 13 cifre fără `RO` (ex. `2950101400123`) | Persoană fizică consumator | **B2C, CNP anonimizat la `0000000000000`** (GDPR) |
| CNP cu `RO` (ex. `RO2950101400123`) | PFA cu CNP plătitor TVA | Tratat ca B2B plătitor |

### 3. Foile `Facturi_Antet` și `Linii_Facturi`

- Pe `Facturi_Antet` adaugi un rând per factură. Sumele (LineExtensionAmount, VATAmount, TaxInclusiveAmount, PayableAmount, CountLine) se calculează automat prin formule din `Linii_Facturi`.
- Pe `Linii_Facturi` adaugi câte un rând per linie. Sumele linie (LineExt = qty×preț, VAT = LineExt × pct/100, Total = LineExt + VAT) se calculează automat.
- Coloana `Invoice` din linii trebuie să se potrivească exact cu `InvoiceID` din antet.

### 4. Generarea

Apasă `Alt + F8`, alege macroul `Genereaza_eFactura_XML_Excel`, apasă `Run`.

Fișierele XML apar în subfolderul `xml/` lângă workbook (ex: `xml/12345678_F2026-001.xml`).

---

## Storno

Factura de stornare se face cu **același tip 380** și **cantitate negativă** pe linie:

| InvoicedQuantity | PriceAmount | LineExtensionAmount | rezultat |
|---|---|---|---|
| `-10` | `5` | `-50` | storno parțial pentru 10 unități la 5 RON |

Nu folosi prețuri negative. Antetul va calcula automat un `PayableAmount` negativ.

---

## Validare la ANAF

Înainte de a încărca în SPV, validează fiecare XML pe pagina ANAF:

**https://www.anaf.ro/uploadxml/**

Dacă apare eroare, deschide XML-ul, fă corecturile în Excel și regenerează.

---

## Limitări v1

- Doar moneda RON.
- Doar tip factură 380 (cu storno prin cantitate negativă).
- Cote TVA suportate: 0, 5, 11, 21 (categoria S; plus categoria O pentru neplătitori).
- Allowance / Charge / Reduceri pe linie nu sunt încă suportate la nivel UI (poți edita manual XML-ul rezultat dacă ai nevoie).
- Aplicația funcționează cu Excel pe Windows. LibreOffice / Mac nu sunt testate (folosește `ADODB.Stream` pentru UTF-8 fără BOM, specific Windows).

## Pe roadmap

- [ ] Suport LibreOffice (înlocuire ADODB.Stream cu API portabil)
- [ ] Tip 384 (factură corectivă), tip 389 (auto-factură)
- [ ] Allowance / Charge la nivel de linie și de antet
- [ ] Buton pe foaia `Citeste-ma` legat direct de macro

Pull requests sunt binevenite.

---

## Cum să contribui

1. Fork repository-ul.
2. Crează un branch: `git checkout -b feat/<numele-feature-ului>`.
3. Modifică `generare_eFactura.bas`. Pentru schimbări de structură ale workbook-ului, modifică și constantele coloanelor de la începutul fișierului `.bas`.
4. Testează generarea pe câteva cazuri reale și verifică pe validatorul ANAF.
5. Deschide un Pull Request.

---

## Disclaimer

Această aplicație este oferită „așa cum este", fără garanții. Nu sunt consultant fiscal — verifică întotdeauna cu propriul contabil înainte de a depinde de output în producție. Autorul nu răspunde pentru erori de raportare sau penalități ANAF rezultate din folosirea acestei aplicații.

---

## Licență

Distribuit sub [Licența MIT](LICENSE). Liber pentru uz comercial și non-comercial.

```
Copyright (c) 2026 <numele tău>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Resurse utile

- Validator ANAF: https://www.anaf.ro/uploadxml/
- ANAF — Despre RO e-Factura: https://www.anaf.ro/anaf/internet/ANAF/despre_anaf/strategii_anaf/proiecte_digitalizare/e.factura
- Specificație CIUS-RO: https://mfinante.gov.ro/static/10/Mfp/efactura/CIUS-RO.pdf
- UBL 2.1 Invoice schema: http://docs.oasis-open.org/ubl/os-UBL-2.1/UBL-2.1.html
