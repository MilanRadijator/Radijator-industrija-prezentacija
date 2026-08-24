from __future__ import annotations

import html
import os
import re
import shutil
import zipfile
from collections import deque
from pathlib import Path

import numpy as np
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageFilter


ROOT = Path(__file__).resolve().parents[1]


def resolve_docx() -> Path:
    override = os.environ.get("RADIJATOR_CATALOG_DOCX")
    candidates = [
        Path(override) if override else None,
        Path.home() / "Downloads" / "KATALOG INDUSTRIJSKIH KOTLOVA.docx",
        Path(r"Z:\02_Konstrukcija\Tijana Vujičić\KATALOG ZA INDUSTRIJSKE KOTLOVE\KATALOG INDUSTRIJSKIH KOTLOVA.docx"),
        Path(r"D:\Prezentacija nikola\KATALOG INDUSTRIJSKIH KOTLOVA.docx"),
    ]
    for candidate in candidates:
        if candidate is not None and candidate.is_file():
            return candidate
    raise FileNotFoundError("KATALOG INDUSTRIJSKIH KOTLOVA.docx nije pronađen.")


DOCX = resolve_docx()
DOCS = ROOT / "docs"
ASSET_DIR = DOCS / "assets" / "full-catalog"
OUT_HTML = DOCS / "index.html"
ALIAS_HTML = DOCS / "full-catalog.html"
RO_HTML = DOCS / "index-ro.html"
RO_ALIAS_HTML = DOCS / "full-catalog-ro.html"
PDF_DOWNLOAD_ENABLED = True
EDITORIAL_ASSET_DIR = DOCS / "assets" / "editorial"
SUPPLEMENTAL_IMAGE_SOURCES = {
    "valvola-tkan-150-presek.png": Path.home() / "Downloads" / "Valvola TKAN 150 - PRESEK.PNG",
    "valvola-tkan-150.png": Path.home() / "Downloads" / "Valvola TKAN 150.PNG",
    "valvola-tkan-300-integra-presek.png": Path.home()
    / "Downloads"
    / "Valvola - TKAN 300 Integra - presek.PNG",
    "valvola-tkan-300-integra.png": Path.home()
    / "Downloads"
    / "Valvola - TKAN 300 Integra.PNG",
    "multiciklon-tkan-300.png": Path.home() / "Downloads" / "Sivi ciklon TKAN 300.PNG",
}

LANGUAGE_CONFIG = {
    "sr": {
        "html_lang": "sr",
        "html_path": OUT_HTML,
        "alias_path": ALIAS_HTML,
        "pdf_filename": "radijator-industrijski-kotlovi.pdf",
        "page_title": "Kompletan katalog | Radijator Inzenjering",
        "aria_main": "Glavna navigacija",
        "brand_title": "Industrijski katalog",
        "brand_aria": "Radijator Inženjering - vrh kataloga",
        "nav_boilers": "Kotlovi",
        "nav_systems": "Sistemi",
        "nav_equipment": "Oprema",
        "nav_contact": "Kontakt",
        "pdf_label": "PDF katalog",
        "hero_topline": "Industrijska termoenergetska rešenja",
        "hero_eyebrow": "Kompletan proizvodni katalog / 2026",
        "hero_title": "Industrijski kotlovi <em>na biomasu</em>",
        "hero_lead": (
            "Pouzdani sistemi visokih snaga, projektovani za efikasnost, "
            "dug radni vek i potpunu kontrolu procesa sagorevanja."
        ),
        "hero_primary": "Pregledaj katalog",
        "hero_secondary": "Preuzmi PDF",
        "scroll": "Skrolujte",
        "toc_summary": "Sadržaj kataloga",
        "top_link": "Vrh kataloga",
        "production_toc": "Proizvodnja i standardi",
        "footer_kicker": "Projektovanje / proizvodnja / podrška",
        "footer_title": "Partner za kompletna termoenergetska rešenja.",
        "footer_address": "Živojina Lazića Solunca 6<br />36000 Kraljevo, Srbija",
        "footer_gallery": "Radijator Inženjering u praksi",
        "footer_product": "Industrijski kotlovi na biomasu",
        "footer_top": "Nazad na vrh",
        "dialog_label": "Uvećani tehnički prikaz",
        "dialog_close": "Zatvori uvećani prikaz",
        "dialog_close_text": "Zatvori",
        "badge_alt": "35 godina iskustva - kvalitet bez kompromisa",
        "badge_title": "35 godina iskustva",
    },
    "ro": {
        "html_lang": "ro",
        "html_path": RO_HTML,
        "alias_path": RO_ALIAS_HTML,
        "pdf_filename": "radijator-industrijski-kotlovi-ro.pdf",
        "page_title": "Catalog complet | Radijator Inzenjering",
        "aria_main": "Navigare principală",
        "brand_title": "Catalog industrial",
        "brand_aria": "Radijator Inženjering - începutul catalogului",
        "nav_boilers": "Cazane",
        "nav_systems": "Sisteme",
        "nav_equipment": "Echipamente",
        "nav_contact": "Contact",
        "pdf_label": "Catalog PDF",
        "hero_topline": "Soluții termoenergetice industriale",
        "hero_eyebrow": "Catalog complet de produse / 2026",
        "hero_title": "Cazane industriale <em>pe biomasă</em>",
        "hero_lead": (
            "Sisteme fiabile de putere mare, proiectate pentru eficiență, "
            "durată lungă de viață și control complet al procesului de ardere."
        ),
        "hero_primary": "Vezi catalogul",
        "hero_secondary": "Descarcă PDF",
        "scroll": "Derulați",
        "toc_summary": "Cuprins catalog",
        "top_link": "Început catalog",
        "production_toc": "Producție și standarde",
        "footer_kicker": "Proiectare / producție / suport",
        "footer_title": "Partener pentru soluții termoenergetice complete.",
        "footer_address": "Živojina Lazića Solunca 6<br />36000 Kraljevo, Serbia",
        "footer_gallery": "Radijator Inženjering în practică",
        "footer_product": "Cazane industriale pe biomasă",
        "footer_top": "Înapoi sus",
        "dialog_label": "Vizualizare tehnică mărită",
        "dialog_close": "Închide vizualizarea mărită",
        "dialog_close_text": "Închide",
        "badge_alt": "35 de ani de experiență - calitate fără compromis",
        "badge_title": "35 de ani de experiență",
    },
}

LANGUAGE_LABELS = {"sr": "SR", "ro": "RO"}

RO_EXACT_TRANSLATIONS = {
    "O nama": "Despre noi",
    'Radijator inženjering" d.o.o. u poslovnom smislu je pravni naslednik zanatske radnje ,,Radijatorˮ koja je osnovana 1991. godine, čija je osnovna delatnost bila montaža i održavanje centralnog grejanja. Prvi toplovodni kotao na čvrsto gorivo izradili smo 1985. godine.': (
        "Radijator Inženjering d.o.o. este continuatorul juridic al atelierului artizanal "
        "„Radijator”, fondat în 1991, a cărui activitate principală a fost montajul și "
        "întreținerea sistemelor de încălzire centrală. Primul cazan de apă caldă pe "
        "combustibil solid l-am produs în anul 1985."
    ),
    "Preduzeće u današnjoj formi postoji od 2002. godine, i iz godine u godinu, napreduje velikim koracima, uvek se trudeći da bude prvo u primeni novih tehnologija, kvalitetu proizvoda i osvajanju novih - evropskih tržišta.": (
        "Compania funcționează în forma actuală din anul 2002 și progresează constant, "
        "urmărind să fie printre primele în aplicarea tehnologiilor noi, în calitatea "
        "produselor și în dezvoltarea de noi piețe europene."
    ),
    "Kako smo proširivali i usavršavali proizvodnju tako smo došli do nivoa da se kotlovi prave najsavremenijim svetskim tehnologijama. Iz oblasti sečenja limova izdvajaju se: sečenje laserom, CNC plazma postupak i CNC probijanje. Postupak zavarivanja izvodi se robotski kao i upotrebom automata. Najbolji pokazatelji kvaliteta proizvoda i usluga jeste činjenica da se svake godine proizvodnja povećava.": (
        "Prin extinderea și perfecționarea producției am ajuns la nivelul la care cazanele "
        "sunt fabricate cu tehnologii moderne de clasă mondială. În prelucrarea tablei se "
        "remarcă tăierea laser, procedeele CNC cu plasmă și perforarea CNC. Sudura se "
        "realizează robotic și cu echipamente automate. Cel mai bun indicator al calității "
        "produselor și serviciilor este faptul că producția crește de la an la an."
    ),
    'Danas "Radijator-inženjering" zapošljava preko 350 radnika od kojih je 40 dipl.maš.ing. koji svakodnevno rade na usavršavanju kvaliteta proizvoda.': (
        "Astăzi, Radijator Inženjering are peste 350 de angajați, dintre care 40 sunt "
        "ingineri mecanici absolvenți care lucrează zilnic la perfecționarea calității "
        "produselor."
    ),
    "Sigurna postojanost kvaliteta, kako proizvoda tako i poslovanja firme, potvrđena je dobijanjem sertifikata sistema kvaliteta ISO 9001:2008.": (
        "Constanța calității produselor și a activității companiei este confirmată prin "
        "certificarea sistemului de management al calității ISO 9001:2008."
    ),
    "Radijator Inženjering je domaći proizvođač kotlova na biomasu sa dugogodišnjom tradicijom, prepoznat po pouzdanim i tehnološki naprednim rešenjima za grejanje. Razvojem sopstvenih proizvodnih procesa i primenom savremenih tehnologija, stvaramo proizvode koji odgovaraju najvišim zahtevima tržišta kada su u pitanju kvalitet, efikasnost i dug vek trajanja.": (
        "Radijator Inženjering este un producător local de cazane pe biomasă cu tradiție "
        "îndelungată, recunoscut pentru soluții de încălzire fiabile și avansate tehnologic. "
        "Prin dezvoltarea propriilor procese de producție și aplicarea tehnologiilor moderne, "
        "realizăm produse care răspund celor mai înalte cerințe ale pieței privind calitatea, "
        "eficiența și durata lungă de exploatare."
    ),
    "Naša proizvodnja obuhvata kotlove snage od 6 do 600 kW, namenjene grejanju porodičnih kuća, stambenih i poslovnih objekata, javnih ustanova i industrijskih postrojenja. Zahvaljujući širokom asortimanu i mogućnosti izrade kaskadnih sistema, u mogućnosti smo da ponudimo optimalno rešenje za svaki objekat i svaku potrebu.": (
        "Producția noastră include cazane cu puteri de la 6 la 600 kW, destinate încălzirii "
        "caselor familiale, clădirilor rezidențiale și comerciale, instituțiilor publice și "
        "instalațiilor industriale. Datorită gamei largi și posibilității realizării "
        "sistemelor în cascadă, putem oferi soluția optimă pentru fiecare obiectiv și fiecare nevoie."
    ),
    "Kvalitet naših proizvoda rezultat je pažljivo odabranih materijala, precizne proizvodnje i rigorozne kontrole kvaliteta u svim fazama procesa. Svaki kotao razvijen je sa ciljem da obezbedi maksimalnu energetsku efikasnost, pouzdan rad i dugoročnu eksploataciju uz minimalne troškove održavanja.": (
        "Calitatea produselor noastre este rezultatul materialelor atent selectate, al "
        "producției precise și al controlului riguros al calității în toate etapele procesului. "
        "Fiecare cazan este dezvoltat pentru a asigura eficiență energetică maximă, funcționare "
        "fiabilă și exploatare pe termen lung cu costuri minime de întreținere."
    ),
    "Posebnu pažnju posvećujemo inovacijama i unapređenju tehnologije proizvodnje, kako bismo korisnicima obezbedili savremena rešenja koja kombinuju visok stepen automatizacije, jednostavno upravljanje i maksimalno iskorišćenje energije.": (
        "Acordăm o atenție specială inovațiilor și îmbunătățirii tehnologiei de producție, "
        "pentru a oferi utilizatorilor soluții moderne care combină un nivel ridicat de "
        "automatizare, operare simplă și valorificare maximă a energiei."
    ),
    "Svi naši proizvodi projektovani su i proizvedeni u skladu sa važećim evropskim standardima i propisima u oblasti kotlova, što predstavlja potvrdu njihove bezbednosti, pouzdanosti i visokog kvaliteta.": (
        "Toate produsele noastre sunt proiectate și fabricate în conformitate cu standardele "
        "și reglementările europene aplicabile în domeniul cazanelor, confirmând siguranța, "
        "fiabilitatea și calitatea lor ridicată."
    ),
    "Izborom Radijator Inženjering kotlova birate domaći proizvod, provereni kvalitet i partnera koji svojim iskustvom, stručnom podrškom i kompletnim sistemskim rešenjima pruža sigurnost tokom celog životnog veka sistema grejanja.": (
        "Alegând cazanele Radijator Inženjering, alegeți un produs local, calitate verificată "
        "și un partener care, prin experiență, suport profesional și soluții complete de sistem, "
        "oferă siguranță pe întreaga durată de viață a sistemului de încălzire."
    ),
    "Kompanija sa 35 godina iskustva u projektovanju, izradi i inovacijama na polju kotlova koji zagrevaju hiljade objekata širom Evrope!": (
        "O companie cu 35 de ani de experiență în proiectarea, fabricarea și inovarea cazanelor "
        "care încălzesc mii de obiective în întreaga Europă!"
    ),
    "Serija TKAN modeli": "Modele seria TKAN",
    "Industrijski kotlovi na biomasu su izrađeni od kotlovskih limova debljine 6 mm i više. Izmenjivač toplote je od bešavnih, kotlovskih cevi. Stepen iskorišćenja preko 90% na pelet. Temperature dimnih gasova na izlazu su od 170 do 190 stepeni pri višim režimima, što uvek možemo da proverimo na displeju automatike. Dostupni su u osegu od 80 – 500 kW.": (
        "Cazanele industriale pe biomasă sunt realizate din tablă de cazan cu grosimea de "
        "6 mm și mai mult. Schimbătorul de căldură este realizat din țevi de cazan fără "
        "sudură. Randamentul depășește 90% la peleți. Temperatura gazelor de ardere la "
        "ieșire este între 170 și 190 de grade în regimuri superioare, valoare care poate "
        "fi verificată pe afișajul automatizării. Sunt disponibile în gama de puteri 60-300 kW."
    ),
    "Industrijski kotlovi na biomasu su izrađeni od kotlovskih limova debljine 6 mm i više. Izmenjivač toplote je od bešavnih, kotlovskih cevi. Stepen iskorišćenja preko 90% na pelet. Temperature dimnih gasova na izlazu su od 170 do 190 stepeni pri višim režimima, što uvek možemo da proverimo na displeju automatike. Dostupni su u opsegu snaga od 60 do 300 kW.": (
        "Cazanele industriale pe biomasă sunt realizate din tablă de cazan cu grosimea de "
        "6 mm și mai mult. Schimbătorul de căldură este realizat din țevi de cazan fără "
        "sudură. Randamentul depășește 90% la peleți. Temperatura gazelor de ardere la "
        "ieșire este între 170 și 190 de grade în regimuri superioare, valoare care poate "
        "fi verificată pe afișajul automatizării. Sunt disponibile în gama de puteri 60-300 kW."
    ),
    "Serija TKAN Integra modeli": "Modele seria TKAN Integra",
    "Industrijski kotao na biomasu predstavlja unapređenu verziju standardnog TKAN kotla. Opremljen je zidanim ložištem, naprednijom automatikom i bogatijom pratećom opremom, čime se postižu veća pouzdanost, bolja efikasnost sagorevanja i šire mogućnosti primene. Izrađen je od kotlovskih limova debljine 6 mm i više, sa izmenjivačem toplote od bešavnih kotlovskih cevi. Stepen iskorišćenja je preko 90% Dostupan je u opsegu snaga od 80 do 500 kW.": (
        "Cazanul industrial pe biomasă reprezintă o versiune îmbunătățită a cazanului "
        "standard TKAN. Este echipat cu focar zidit, automatizare mai avansată și echipamente "
        "auxiliare mai bogate, obținându-se fiabilitate mai mare, eficiență mai bună a arderii "
        "și posibilități mai largi de utilizare. Este realizat din tablă de cazan cu grosimea "
        "de 6 mm și mai mult, cu schimbător de căldură din țevi de cazan fără sudură. "
        "Randamentul depășește 90%, iar gama de puteri disponibilă este 80-600 kW."
    ),
    "Industrijski kotao na biomasu predstavlja unapređenu verziju standardnog TKAN kotla. Opremljen je zidanim ložištem, naprednijom automatikom i bogatijom pratećom opremom, čime se postižu veća pouzdanost, bolja efikasnost sagorevanja i šire mogućnosti primene. Izrađen je od kotlovskih limova debljine 6 mm i više, sa izmenjivačem toplote od bešavnih kotlovskih cevi. Stepen iskorišćenja je preko 90% Dostupan je u opsegu snaga od 80 do 600 kW.": (
        "Cazanul industrial pe biomasă reprezintă o versiune îmbunătățită a cazanului "
        "standard TKAN. Este echipat cu focar zidit, automatizare mai avansată și echipamente "
        "auxiliare mai bogate, obținându-se fiabilitate mai mare, eficiență mai bună a arderii "
        "și posibilități mai largi de utilizare. Este realizat din tablă de cazan cu grosimea "
        "de 6 mm și mai mult, cu schimbător de căldură din țevi de cazan fără sudură. "
        "Randamentul depășește 90%, iar gama de puteri disponibilă este 80-600 kW."
    ),
    "Kaskadni sistemi": "Sisteme în cascadă",
    "Kaskadni sistemi predstavljaju kombinaciju dva ili više kotlova povezanih u jedinstven sistem sa zajedničkim silosom za skladištenje peleta. Ovakva konfiguracija omogućava veću ukupnu instalisanu snagu, pouzdaniji rad, ravnomernu raspodelu opterećenja i veću energetsku efikasnost sistema.": (
        "Sistemele în cascadă reprezintă o combinație de două sau mai multe cazane conectate "
        "într-un sistem unic, cu siloz comun pentru depozitarea peleților. O astfel de "
        "configurație permite o putere totală instalată mai mare, funcționare mai fiabilă, "
        "distribuție uniformă a sarcinii și eficiență energetică mai ridicată a sistemului."
    ),
    "Dodatna oprema": "Echipamente suplimentare",
    "Pored kotlova, u ponudi je kompletna prateća oprema za formiranje funkcionalnog sistema grejanja. Asortiman obuhvata silose za skladištenje peleta, pužne transportere, elevatore, bafer rezervoare, automatiku i ostalu opremu potrebnu za pouzdan transport, doziranje i kontrolu peleta, kao i siguran i efikasan rad celokupnog sistema.": (
        "Pe lângă cazane, oferta include echipamente auxiliare complete pentru formarea unui "
        "sistem funcțional de încălzire. Gama cuprinde silozuri pentru depozitarea peleților, "
        "transportoare melcate, elevatoare, rezervoare tampon, automatizare și alte echipamente "
        "necesare pentru transportul, dozarea și controlul fiabil al peleților, precum și "
        "pentru funcționarea sigură și eficientă a întregului sistem."
    ),
    "Položaj TKAN običnog i TKAN Integra kotla u kotlarnici": (
        "Poziționarea cazanului TKAN și TKAN Integra în camera tehnică"
    ),
    "(U ponudi snaga od 60 – 300 [kW] ) Kotao TKAN je razvijen sa ciljem da RADIJATOR INŽENJERING ponudi tržištu kotao koji je po svojim mehaničkim i termičkim osobinama izrazito namenjen biomasi kao gorivu. Sa druge strane zahtevi tržišta su uvek okrenuti ka što većoj univerzalnosti goriva, tako da je TKAN moguće ložiti i sa drvetom i tada je loženje ručno.": (
        "(Disponibil în gama de puteri 60 - 300 [kW]) Cazanul TKAN a fost dezvoltat cu scopul "
        "ca RADIJATOR INŽENJERING să ofere pieței un cazan ale cărui proprietăți mecanice și "
        "termice sunt adaptate în mod special biomasei ca combustibil. Pe de altă parte, "
        "cerințele pieței sunt orientate către o universalitate cât mai mare a combustibilului, "
        "astfel încât TKAN poate fi alimentat și cu lemn, caz în care alimentarea este manuală."
    ),
    "Po spoljašnjem dizajnu, dimenzijama ložišta, otvorima za loženje i čišćenje TKAN je zadržao sve dobre osobine predhodnih modela po kojima je RADIJATOR INŽENJERING prepoznatljiv na tržištu. Vodeni deo kotla, njegov način izmene toplote između dimnih gasova i vode putem cevnog izmenjivača, prilagođen je biomasi. Zbog upotrebe ventilatora tj. prinudne promaje put dimnih gasova duži je nego kod standardnih kotlova. Iz istih razloga moguća je primena usmerivača dimnih gasova tzv. turbulatora koji dodatno povećavaju stepen iskorišćenja kotla. Turbulatori su spirale napravljene od specijalnog materijala.": (
        "Prin designul exterior, dimensiunile focarului și deschiderile pentru alimentare și "
        "curățare, TKAN păstrează toate calitățile modelelor anterioare prin care RADIJATOR "
        "INŽENJERING este recunoscut pe piață. Partea de apă a cazanului și transferul de "
        "căldură dintre gazele de ardere și apă prin schimbătorul tubular sunt adaptate "
        "biomasei. Datorită utilizării ventilatorului, respectiv a tirajului forțat, traseul "
        "gazelor de ardere este mai lung decât la cazanele standard. Din același motiv se pot "
        "utiliza dirijori ai gazelor de ardere, așa-numitele turbultoare, care cresc suplimentar "
        "randamentul cazanului. Turbulatoarele sunt spirale realizate din material special."
    ),
    "Stepen korisnosti na pelet je preko 90%. Pri normalnim režimima temperatura dimnih gasova na izlazu je oko 160  ̊ C, a pri maksimalnim režimima je ispod 180  ̊ C. Ove vrednosti mogu u svakom trenutku da se očitaju na displeju.": (
        "Randamentul la peleți depășește 90%. În regimuri normale, temperatura gazelor de "
        "ardere la ieșire este de aproximativ 160 °C, iar în regimuri maxime este sub 180 °C. "
        "Aceste valori pot fi citite în orice moment pe afișaj."
    ),
    "Svi delovi vodenog dela kotla izrađeni su od bešavnih cevi kvaliteta ST 35.4 i kotlovskih limova debljine 5mm i više, u zavisnosti od snage kotla. Limovi su kvaliteta 1.0425 EU standard odnosno P265GH standard EUII. Ložište je po svom principu rada tzv. „izviruće“, gde gorivo iz zone transporta ide vertikalno uvis tj. izvire do zone sagorevanja. Napravljeno je od masivnih izolacijskih materijala i sivog liva. Transport goriva obezbeđen je pužnim transporterima.": (
        "Toate componentele părții de apă a cazanului sunt realizate din țevi fără sudură de "
        "calitate ST 35.4 și din tablă de cazan cu grosimea de 5 mm și mai mult, în funcție "
        "de puterea cazanului. Tablele sunt de calitate 1.0425 conform standardului UE, "
        "respectiv P265GH conform standardului EUII. Focarul funcționează pe principiul "
        "așa-numitului focar ascendent, în care combustibilul din zona de transport se deplasează "
        "vertical în sus către zona de ardere. Este realizat din materiale izolante masive și "
        "fontă cenușie. Transportul combustibilului este asigurat prin transportoare melcate."
    ),
    "Tabela 1. Raspoložive snage": "Tabelul 1. Puteri disponibile",
    "Primena TKAN kotlova u kaskadi": "Utilizarea cazanelor TKAN în cascadă",
    "Specifičnosti TKAN kaskadnih sistema": "Particularitățile sistemelor TKAN în cascadă",
    "Kompanija Radijator Inženjering razvila je seriju kotlova TKAN prvenstvenstveno za sagorevanje biomase (pelet, koštica voća) i drveta. Kada se ovi kotlovi povezuju u kaskadne sisteme, dobija se izuzetno moćno i fleksibilno rešenje za grejanje velikih objekata.": (
        "Compania Radijator Inženjering a dezvoltat seria de cazane TKAN în principal pentru "
        "arderea biomasei (peleți, sâmburi de fructe) și a lemnului. Atunci când aceste cazane "
        "sunt conectate în sisteme în cascadă, se obține o soluție extrem de puternică și "
        "flexibilă pentru încălzirea obiectivelor mari."
    ),
    "Za kaskadne sisteme najčešće se koriste industrijski modeli veće snage, kao što su TKAN 100, 150, 200, 250 i 300 kW.": (
        "Pentru sistemele în cascadă se utilizează cel mai frecvent modele industriale de "
        "putere mai mare, precum TKAN 100, 150, 200, 250 și 300 kW."
    ),
    "Pokrivanje velikih snaga: Povezivanjem npr. dva kotla TKAN 300 u kaskadu, dobija se sistem ukupne snage od 600 kW koji može da greje hotele, proizvodne hale ili stambene komplekse.": (
        "Acoperirea puterilor mari: prin conectarea, de exemplu, a două cazane TKAN 300 în "
        "cascadă, se obține un sistem cu putere totală de 600 kW, capabil să încălzească "
        "hoteluri, hale de producție sau ansambluri rezidențiale."
    ),
    "Modularnost i fleksibilnost: U prelaznim periodima (jesen/proleće) radi samo jedan TKAN kotao na optimalnom režimu, dok se drugi pali tek kada spoljna temperatura drastično padne.": (
        "Modularitate și flexibilitate: în perioadele de tranziție (toamnă/primăvară) "
        "funcționează un singur cazan TKAN în regim optim, iar al doilea pornește doar când "
        "temperatura exterioară scade semnificativ."
    ),
    "Upravljanje i automatika: TKAN kotlovi poseduju naprednu elektroniku koja preko spoljnih kaskadnih regulatora omogućava sinhronizovan rad. Automatika prati temperaturu u hidrauličnoj skretnici i komanduje koji će kotao startovati.": (
        "Comandă și automatizare: cazanele TKAN dispun de electronică avansată care, prin "
        "regulatoare externe de cascadă, permite funcționarea sincronizată. Automatizarea "
        "urmărește temperatura în separatorul hidraulic și comandă cazanul care trebuie să pornească."
    ),
    "Kontinuirano snabdevanje gorivom: Industrijski TKAN kotlovi dolaze sa dnevnim silozima (npr. 800 litara) koji se preko dodatnih pužnih transportera mogu povezati sa jednim velikim, centralnim silosom za pelet koji snabdeva celu kaskadu.": (
        "Alimentare continuă cu combustibil: cazanele industriale TKAN sunt livrate cu silozuri "
        "zilnice (de exemplu 800 litri), care pot fi conectate prin transportoare melcate "
        "suplimentare la un siloz central mare pentru peleți, ce alimentează întreaga cascadă."
    ),
    "Sigurnost i kontinuitet: Ukoliko je na jednom kotlu potrebno uraditi čišćenje pepela ili redovan servis, hidraulički sistem i kaskadna automatika omogućavaju da drugi kotao nesmetano nastavi rad, tako da objekat nikada ne ostaje bez grejanja.": (
        "Siguranță și continuitate: dacă la un cazan este necesară curățarea cenușii sau "
        "service-ul periodic, sistemul hidraulic și automatizarea în cascadă permit celuilalt "
        "cazan să continue funcționarea fără întreruperi, astfel încât obiectivul să nu rămână "
        "niciodată fără încălzire."
    ),
    "PROIZVODNI PROGRAM – INDUSTRIJA": "PROGRAM DE PRODUCȚIE - INDUSTRIE",
    "PROIZVODNI PROGRAM - INDUSTRIJA": "PROGRAM DE PRODUCȚIE - INDUSTRIE",
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM TKAN MODEL": (
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ - MODEL TKAN"
    ),
    "PRESEK TKAN KOTLA SA OPISOM ELEMENATA": (
        "SECȚIUNE CAZAN TKAN CU DESCRIEREA ELEMENTELOR"
    ),
    "TABELA SA DIMENZIJAMA TKAN KOTLA": "TABEL CU DIMENSIUNILE CAZANULUI TKAN",
    "TABELA SA DIMENZIJAMA TKAN SILOSA": "TABEL CU DIMENSIUNILE SILOZULUI TKAN",
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA MODEL": (
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ, DESPRĂFUIRE ȘI CICLON - MODEL TKAN INTEGRA"
    ),
    "PRESEK TKAN INTEGRA KOTLA SA OPISOM ELEMENATA": (
        "SECȚIUNE CAZAN TKAN INTEGRA CU DESCRIEREA ELEMENTELOR"
    ),
    "TABELA SA DIMENZIJAMA TKAN INTEGRA SILOSA": (
        "TABEL CU DIMENSIUNILE SILOZULUI TKAN INTEGRA"
    ),
    "POLOŽAJ TKAN OBIČNOG I TKAN INTEGRA KOTLA U KOTLARNICI": (
        "POZIȚIONAREA CAZANULUI TKAN ȘI TKAN INTEGRA ÎN CAMERA TEHNICĂ"
    ),
    "POLOŽAJ TKAN OBIČNOG i TKAN INTEGRA KOTLA U KOTLARNICI": (
        "POZIȚIONAREA CAZANULUI TKAN ȘI TKAN INTEGRA ÎN CAMERA TEHNICĂ"
    ),
    "KASKADNI SISTEMI": "SISTEME ÎN CASCADĂ",
    "DODATNA OPREMA": "ECHIPAMENTE SUPLIMENTARE",
    "AUTOMATSKI TRANSPORT PELETA": "TRANSPORT AUTOMAT AL PELEȚILOR",
    "SERIJA TKAN MODELI": "SERIA DE MODELE TKAN",
    "SERIJA TKAN INTEGRA MODELI": "SERIA DE MODELE TKAN INTEGRA",
    "DIMENZIJE": "DIMENSIUNI",
    "TIP KOTLA": "TIP CAZAN",
    "Tip kotla": "Tip cazan",
    "Jed.mere": "UM",
    "Snaga": "Putere",
    "Radni pritisak": "Presiune de lucru",
    "Probni pritisak": "Presiune de probă",
    "Zapremina vode u kotlu": "Volum apă în cazan",
    "Masa kotla": "Masă cazan",
    "Masa silosa": "Masă siloz",
    "Količina peleta koja staje u silosu": "Cantitate de peleți în siloz",
    "Turbulatori": "Turbulatoare",
    "Cevni izmenjivač": "Schimbător tubular",
    "Vrata za čišćenje cevnog izmenjivača i samog kotla": (
        "Ușă pentru curățarea schimbătorului tubular și a cazanului"
    ),
    "Razvodni ormar sa automatizacije": "Dulap de comandă cu automatizare",
    "Razvodni ormar sa automatikom": "Dulap de comandă cu automatizare",
    "Razvodni ormar sa automatizare": "Dulap de comandă cu automatizare",
    "Kanta za pepeo": "Cutie pentru cenușă",
    "Vrata za loženje i potpalu": "Ușă pentru alimentare și aprindere",
    "Spirala za automat izbacivanje pepela iz prostora ložišta": (
        "Spirală pentru evacuarea automată a cenușii din zona focarului"
    ),
    "Spirala za automatsko izbacivanje pepela iz prostora ložišta": (
        "Spirală pentru evacuarea automată a cenușii din zona focarului"
    ),
    "Ložište kotla": "Focarul cazanului",
    "Liveni segmenti": "Segmente turnate",
    "Motor za pokretanje spirale za automat čišćenje ložišta": (
        "Motor pentru acționarea spiralei de curățare automată a focarului"
    ),
    "Motor za pokretanje spirale za automatsko čišćenje ložišta": (
        "Motor pentru acționarea spiralei de curățare automată a focarului"
    ),
    "Donja osovina pužnog transportera": "Ax inferior al transportorului melcat",
    "Ćelijski dozator (valvola)": "Dozator celular (valvola)",
    "Gornja osovina pužnog transportera": "Ax superior al transportorului melcat",
    "Izmenjivač termičkog osiguranja": "Schimbător pentru protecție termică",
    "Boca komprimovanog vazduha": "Rezervor de aer comprimat",
    "Impulsni elektroventil": "Electrovalvă cu impuls",
    "Ozida ložišta": "Căptușeala focarului",
    "Ožid ložišta": "Căptușeala focarului",
    "Centrifugalni ventilator multiciklona": "Ventilator centrifugal al multiciclonului",
    "Multiciklon": "Multiciclon",
    "(U ponudi snaga od 80 – 500 [kW] )": "(Disponibil în gama de puteri 80 - 500 [kW])",
    "(U ponudi snaga od 80 - 500 [kW] )": "(Disponibil în gama de puteri 80 - 500 [kW])",
    "(U ponudi snaga od 80 – 600 [kW] )": "(Disponibil în gama de puteri 80 - 600 [kW])",
    "(U ponudi snaga od 80 - 600 [kW] )": "(Disponibil în gama de puteri 80 - 600 [kW])",
    "Pozicioniranje kotla u kotlarnici": "Poziționarea cazanului în camera tehnică",
    "Kotlarnica mora biti obezbeđena od smrzavanja. Podloga za kotao u kotlarnici mora biti od nezapaljivog materijala. Preporučene vrednosti udaljenosti sve četiri strane kotla u odnosu na zidove kotlarnice ili neka druga kruta tela (akomulacioni bojler itd.) prikazane su tablično slici ispod.": (
        "Camera tehnică trebuie protejată împotriva înghețului. Baza pe care se amplasează "
        "cazanul trebuie să fie realizată din material incombustibil. Valorile recomandate "
        "ale distanțelor pe toate cele patru laturi ale cazanului față de pereții camerei "
        "tehnice sau față de alte corpuri rigide (rezervor de acumulare etc.) sunt prezentate "
        "în tabelul și imaginea de mai jos."
    ),
    "VEĆI DNEVNI SILOS": "SILOZ ZILNIC DE CAPACITATE MAI MARE",
    "Standardni TKAN kotlovi opremljeni su dnevnim silosom, dok proizvođač za veće sisteme omogućava izradu većih spoljnih silosa sa posebnim dimenzijama. U zavisnosti od potreba sistema, mogu se izraditi silosi kapaciteta nekoliko desetina tona sa kofičastim elevatorom.": (
        "Cazanele standard TKAN sunt echipate cu siloz zilnic, iar pentru sisteme mai mari "
        "producătorul permite realizarea unor silozuri externe de capacitate mai mare, cu "
        "dimensiuni speciale. În funcție de necesarul sistemului, se pot realiza silozuri cu "
        "capacitate de câteva zeci de tone, cu elevator cu cupe."
    ),
    "Veliki spoljni silos povezuje se sa dnevnim silosom kotla putem pužnih transportera, čime se omogućava automatsko dopremanje dnevnog silosa putem sondi minimuma i maksimuma. Za skladištenje većih količina peleta mogu se koristiti i Jumbo vreće.": (
        "Silozul exterior mare se conectează la silozul zilnic al cazanului prin transportoare "
        "melcate, ceea ce permite alimentarea automată a silozului zilnic prin sonde de minim "
        "și maxim. Pentru depozitarea unor cantități mai mari de peleți se pot utiliza și saci Jumbo."
    ),
    "Proizvođač predviđa: pužne transportere, pogone/motore pužnih transportera, povezivanje velikog i dnevnog silosa, automatsko dopunjavanje dnevnog silosa, sonde minimuma i maksimuma u dnevnom silosu.": (
        "Producătorul prevede transportoare melcate, acționări/motoare pentru transportoare, "
        "conectarea silozului mare cu silozul zilnic, completarea automată a silozului zilnic "
        "și sonde de minim și maxim în silozul zilnic."
    ),
    "Kod velikih sistema preporučuje se veliki silos sa kofičastim elevatorom, pužnim transporterima i automatskim dopunjavanjem dnevnog silosa putem sondi minimuma i maksimuma. Za skladištenje većih količina peleta mogu se koristiti i Jumbo vreće.": (
        "La sistemele mari se recomandă un siloz mare cu elevator cu cupe, transportoare melcate "
        "și completare automată a silozului zilnic prin sonde de minim și maxim. Pentru depozitarea "
        "unor cantități mai mari de peleți se pot utiliza și saci Jumbo."
    ),
    "veliki centralni siloz → pužni transporter → dnevni siloz TKAN → puž do ložišta": (
        "siloz central mare -> transportor melcat -> siloz zilnic TKAN -> melc către focar"
    ),
    "veliki centralni silos → pužni transporter → dnevni silos TKAN → puž do ložišta": (
        "siloz central mare -> transportor melcat -> siloz zilnic TKAN -> melc către focar"
    ),
}

RO_PHRASE_TRANSLATIONS = [
    ("kotlove snage od 15 do 500 kW", "cazane cu puteri de la 6 la 600 kW"),
    ("kotlove snage od 6 do 600 kW", "cazane cu puteri de la 6 la 600 kW"),
    ("snage od 15 do 500 kW", "puteri de la 6 la 600 kW"),
    ("snage od 6 do 600 kW", "puteri de la 6 la 600 kW"),
    ("Radijator inženjering", "Radijator Inženjering"),
    ("Radijator Inženjering", "Radijator Inženjering"),
    ("industrijski kotao na biomasu", "cazan industrial pe biomasă"),
    ("Industrijski kotao na biomasu", "Cazan industrial pe biomasă"),
    ("industrijski kotlovi na biomasu", "cazane industriale pe biomasă"),
    ("Industrijski kotlovi na biomasu", "Cazane industriale pe biomasă"),
    ("kotlovi na biomasu", "cazane pe biomasă"),
    ("kotlovskih limova debljine 6 mm i više", "tablă de cazan cu grosimea de 6 mm și mai mult"),
    ("kotlovskih cevi", "țevi de cazan"),
    ("izmenjivačem toplote", "schimbător de căldură"),
    ("izmenjivač toplote", "schimbător de căldură"),
    ("stepen iskorišćenja", "randament"),
    ("temperatura dimnih gasova", "temperatura gazelor de ardere"),
    ("displeju automatike", "afișajul automatizării"),
    ("dostupan je u opsegu snaga", "este disponibil în gama de puteri"),
    ("opsegu snaga", "gama de puteri"),
    ("pelet", "peleți"),
    ("peleta", "peleți"),
    ("silosa", "silozului"),
    ("silos", "siloz"),
    ("ložištem", "focar"),
    ("ložište", "focar"),
    ("automatikom", "automatizare"),
    ("automatika", "automatizare"),
    ("automatsko", "automat"),
    ("automatskim", "automat"),
    ("sagorevanja", "arderii"),
    ("pouzdanost", "fiabilitate"),
    ("efikasnost", "eficiență"),
    ("primene", "utilizare"),
    ("proizvodnja", "producție"),
    ("proizvodnju", "producția"),
    ("proizvodnih procesa", "proceselor de producție"),
    ("kvalitet", "calitate"),
    ("tržišta", "pieței"),
    ("tržište", "piață"),
    ("evropskih", "europene"),
    ("savremenih tehnologija", "tehnologii moderne"),
    ("lasersko sečenje", "tăiere laser"),
    ("CNC plazma postupak", "procedeu CNC plasmă"),
    ("CNC probijanje", "perforare CNC"),
    ("robotsko zavarivanje", "sudură robotică"),
    ("zavarivanje automatima", "sudură automatizată"),
    ("zaposljava preko 350 radnika", "are peste 350 de angajați"),
    ("preko 350 radnika", "peste 350 de angajați"),
    ("40 diplomiranih masinskih inzenjera", "40 de ingineri mecanici absolvenți"),
    ("dipl. masinskih inzenjera", "ingineri mecanici absolvenți"),
    ("zaposlenih", "angajați"),
    ("izvoz u 27+ zemalja EU", "export în 27+ țări UE"),
    ("Tehnologija i kvalitet", "Tehnologie și calitate"),
    ("Proizvodnja po savremenim evropskim standardima", "Producție conform standardelor europene moderne"),
    ("Kako se proizvodnja sirila i usavrsavala", "Pe măsură ce producția s-a extins și s-a perfecționat"),
    ("Danas Radijator Inzenjering", "Astăzi Radijator Inženjering"),
    ("unapredjenju kvaliteta proizvoda", "îmbunătățirea calității produselor"),
    ("opremljen", "echipat"),
    ("opremljeni", "echipate"),
    ("standardno", "standard"),
    ("model", "model"),
    ("Tabela", "Tabel"),
    ("tabela", "tabel"),
    ("dimenzijama", "dimensiunile"),
    ("Dimenzije", "Dimensiuni"),
    ("Masa", "Masă"),
    ("Radni", "De lucru"),
    ("Probni", "De probă"),
]

RO_PREFIX_TRANSLATIONS = [
    (
        "Radijator inženjering",
        "Radijator Inženjering d.o.o. este continuatorul juridic al atelierului artizanal "
        "„Radijator”, fondat în 1991, a cărui activitate principală a fost montajul și "
        "întreținerea sistemelor de încălzire centrală. Primul cazan de apă caldă pe "
        "combustibil solid l-am produs în anul 1985.",
    ),
    (
        "Preduzeće u današnjoj formi postoji od 2002.",
        "Compania funcționează în forma actuală din anul 2002 și progresează constant, "
        "urmărind să fie printre primele în aplicarea tehnologiilor noi, în calitatea "
        "produselor și în dezvoltarea de noi piețe europene.",
    ),
    (
        "Kako smo proširivali i usavršavali proizvodnju",
        "Prin extinderea și perfecționarea producției am ajuns la nivelul la care cazanele "
        "sunt fabricate cu tehnologii moderne de clasă mondială. În prelucrarea tablei se "
        "remarcă tăierea laser, procedeele CNC cu plasmă și perforarea CNC. Sudura se "
        "realizează robotic și cu echipamente automate. Cel mai bun indicator al calității "
        "produselor și serviciilor este faptul că producția crește de la an la an.",
    ),
    (
        'Danas "Radijator-inženjering"',
        "Astăzi, Radijator Inženjering are peste 350 de angajați, dintre care 40 sunt "
        "ingineri mecanici absolvenți care lucrează zilnic la perfecționarea calității "
        "produselor.",
    ),
    (
        "Sigurna postojanost kvaliteta",
        "Constanța calității produselor și a activității companiei este confirmată prin "
        "certificarea sistemului de management al calității ISO 9001:2008.",
    ),
    (
        "(U ponudi snaga od 60",
        "(Disponibil în gama de puteri 60 - 300 [kW]) Cazanul TKAN a fost dezvoltat cu scopul "
        "ca RADIJATOR INŽENJERING să ofere pieței un cazan ale cărui proprietăți mecanice și "
        "termice sunt adaptate în mod special biomasei ca combustibil. Pe de altă parte, "
        "cerințele pieței sunt orientate către o universalitate cât mai mare a combustibilului, "
        "astfel încât TKAN poate fi alimentat și cu lemn, caz în care alimentarea este manuală.",
    ),
    (
        "Po spoljašnjem dizajnu",
        "Prin designul exterior, dimensiunile focarului și deschiderile pentru alimentare și "
        "curățare, TKAN păstrează toate calitățile modelelor anterioare prin care RADIJATOR "
        "INŽENJERING este recunoscut pe piață. Partea de apă a cazanului și transferul de "
        "căldură dintre gazele de ardere și apă prin schimbătorul tubular sunt adaptate "
        "biomasei. Datorită utilizării ventilatorului, respectiv a tirajului forțat, traseul "
        "gazelor de ardere este mai lung decât la cazanele standard. Din același motiv se pot "
        "utiliza dirijori ai gazelor de ardere, așa-numitele turbulatoare, care cresc suplimentar "
        "randamentul cazanului. Turbulatoarele sunt spirale realizate din material special.",
    ),
    (
        "Stepen korisnosti na pelet",
        "Randamentul la peleți depășește 90%. În regimuri normale, temperatura gazelor de "
        "ardere la ieșire este de aproximativ 160 °C, iar în regimuri maxime este sub 180 °C. "
        "Aceste valori pot fi citite în orice moment pe afișaj.",
    ),
    (
        "Svi delovi vodenog dela kotla",
        "Toate componentele părții de apă a cazanului sunt realizate din țevi fără sudură de "
        "calitate ST 35.4 și din tablă de cazan cu grosimea de 5 mm și mai mult, în funcție "
        "de puterea cazanului. Tablele sunt de calitate 1.0425 conform standardului UE, "
        "respectiv P265GH conform standardului EUII. Focarul funcționează pe principiul "
        "așa-numitului focar ascendent, în care combustibilul din zona de transport se deplasează "
        "vertical în sus către zona de ardere. Este realizat din materiale izolante masive și "
        "fontă cenușie. Transportul combustibilului este asigurat prin transportoare melcate.",
    ),
    (
        "Kompanija Radijator Inženjering razvila je seriju kotlova TKAN",
        "Compania Radijator Inženjering a dezvoltat seria de cazane TKAN în principal pentru "
        "arderea biomasei (peleți, sâmburi de fructe) și a lemnului. Atunci când aceste cazane "
        "sunt conectate în sisteme în cascadă, se obține o soluție extrem de puternică și "
        "flexibilă pentru încălzirea obiectivelor mari.",
    ),
    (
        "Za kaskadne sisteme najčešće",
        "Pentru sistemele în cascadă se utilizează cel mai frecvent modele industriale de "
        "putere mai mare, precum TKAN 100, 150, 200, 250 și 300 kW.",
    ),
    (
        "Pokrivanje velikih snaga",
        "Acoperirea puterilor mari: prin conectarea, de exemplu, a două cazane TKAN 300 în "
        "cascadă, se obține un sistem cu putere totală de 600 kW, capabil să încălzească "
        "hoteluri, hale de producție sau ansambluri rezidențiale.",
    ),
    (
        "Modularnost i fleksibilnost",
        "Modularitate și flexibilitate: în perioadele de tranziție (toamnă/primăvară) "
        "funcționează un singur cazan TKAN în regim optim, iar al doilea pornește doar când "
        "temperatura exterioară scade semnificativ.",
    ),
    (
        "Upravljanje i automatika",
        "Comandă și automatizare: cazanele TKAN dispun de electronică avansată care, prin "
        "regulatoare externe de cascadă, permite funcționarea sincronizată. Automatizarea "
        "urmărește temperatura în separatorul hidraulic și comandă cazanul care trebuie să pornească.",
    ),
    (
        "Kontinuirano snabdevanje gorivom",
        "Alimentare continuă cu combustibil: cazanele industriale TKAN sunt livrate cu silozuri "
        "zilnice, care pot fi conectate prin transportoare melcate suplimentare la un siloz "
        "central mare pentru peleți, ce alimentează întreaga cascadă.",
    ),
    (
        "Sigurnost i kontinuitet",
        "Siguranță și continuitate: dacă la un cazan este necesară curățarea cenușii sau "
        "service-ul periodic, sistemul hidraulic și automatizarea în cascadă permit celuilalt "
        "cazan să continue funcționarea fără întreruperi, astfel încât obiectivul să nu rămână "
        "niciodată fără încălzire.",
    ),
    (
        "TKAN INTEGRA predstavlja novu generaciju",
        "TKAN INTEGRA reprezintă o nouă generație de cazane industriale pe biomasă, "
        "dezvoltată ca îmbunătățire a modelului standard TKAN. Modelul a fost creat ca "
        "răspuns la cerințele pieței pentru un nivel mai ridicat de automatizare, eficiență "
        "energetică superioară și fiabilitate mai mare în exploatare, păstrând soluțiile "
        "constructive verificate pentru care RADIJATOR INŽENJERING este recunoscut."
    ),
    (
        "Zahvaljujući zidanom ložištu",
        "Datorită focarului zidit, sistemului de ardere îmbunătățit, automatizării moderne "
        "și echipării standard mai bogate, TKAN INTEGRA asigură funcționare stabilă, "
        "valorificare maximă a energiei și durată lungă de viață a componentelor cheie. "
        "Construcția cazanului este realizată din tablă de cazan cu grosimea de 6 mm și "
        "mai mult, iar schimbătorul tubular este realizat din țevi de cazan fără sudură, "
        "asigurând rezistență mecanică ridicată, fiabilitate și durabilitate."
    ),
    (
        "Zidano ložište predstavlja jednu od ključnih",
        "Focarul zidit este unul dintre avantajele principale ale modelului TKAN INTEGRA. "
        "Această construcție permite arderea completă și stabilă a combustibilului, emisii "
        "minime de gaze nocive și particule de praf, precum și utilizarea maximă a energiei "
        "conținute în peleți. În același timp, piesele din oțel ale cazanului nu sunt expuse "
        "direct flăcării, ceea ce prelungește semnificativ durata de viață a cazanului."
    ),
    (
        "Za smanjenje količine čestica koje odlaze u dimnjak",
        "Pentru reducerea cantității de particule evacuate către coș este prevăzut un ciclon, "
        "iar la configurațiile mai mari un multiciclon cu ventilator centrifugal. Producătorul "
        "recomandă în mod special ciclonul atunci când se utilizează curățarea pneumatică a "
        "schimbătorului, deoarece atunci cenușa și funinginea sunt evacuate suplimentar din "
        "cazan. La modelele TKAN Integra mai mari se utilizează multiciclon cu ventilator "
        "centrifugal. Mai exact, la modelele TKAN 80 Integra și TKAN 100 Integra ventilatorul "
        "este montat pe racordul de fum, iar la modelele TKAN 150 Integra, TKAN 200 Integra, "
        "TKAN 250 Integra și TKAN 300 Integra este proiectat multiciclon cu ventilator centrifugal."
    ),
    (
        "Kod pojedinih TKAN konfiguracija koristi se ventilator",
        "La anumite configurații TKAN se utilizează ventilator pe partea de evacuare a gazelor "
        "arse, iar la sistemele mai mari multiciclonul poate fi echipat cu ventilator "
        "centrifugal. Acest aspect este important la proiectarea întregului coș de fum, mai "
        "ales în cazul mai multor cazane, traseelor de fum mai lungi, ciclonului/multiciclonului, "
        "numărului mai mare de coturi sau al unui coș comun."
    ),
    (
        "Kotlovi serije TKAN INTEGRA standardno su opremljeni",
        "Cazanele din seria TKAN INTEGRA sunt echipate standard cu sistem pentru curățarea "
        "automată a zonei din jurul focarului, iar schimbătorul tubular se curăță automat cu "
        "aer comprimat. Prin impulsuri periodice de aer, sistemul elimină depunerile de "
        "funingine din țevile de fum, menține randamentul ridicat al cazanului și reduce "
        "semnificativ necesarul de întreținere manuală."
    ),
    (
        "Za pouzdan i efikasan rad industrijskih kotlovskih postrojenja",
        "Pentru funcționarea fiabilă și eficientă a instalațiilor industriale de cazane, "
        "sistemul poate fi echipat cu componente suplimentare adaptate cerințelor obiectivului "
        "și modului de utilizare."
    ),
    (
        "Dodatna oprema obuhvata sisteme za automatsko doziranje",
        "Echipamentele suplimentare includ sisteme pentru dozarea automată și transportul "
        "combustibilului, depozitarea peleților, reglarea automată a funcționării, precum și "
        "echipamente pentru conectarea și controlul în cascadă al cazanelor."
    ),
    (
        "Rešenja se projektuju prema kapacitetu kotlarnice",
        "Soluțiile sunt proiectate în funcție de capacitatea camerei tehnice, autonomia necesară "
        "de funcționare și spațiul disponibil, cu scopul de a obține un nivel ridicat de "
        "automatizare, fiabilitate și consum optim de combustibil."
    ),
    (
        "U ložišnom delu, za automatsko izdvajanje pepela",
        "În zona focarului, pentru evacuarea automată a cenușii, se montează două spirale melcate "
        "cu acționări electrice proprii. Acestea transportă cenușa în două cutii care trebuie "
        "golite periodic."
    ),
    (
        "Na vrata izmenjivačkog sklopa cevi ugrađuje se sistem",
        "Pe ușa ansamblului schimbătorului tubular se montează un sistem de electrovalve care "
        "eliberează periodic aer sub presiune și curăță astfel țevile cazanului de cenușă și "
        "funingine. Este necesară o sursă de aer comprimat cu capacitate adecvată, precum și "
        "automatizare care controlează acest proces."
    ),
    (
        "Zbog smanjene emisije čestica pepela u vazduhu",
        "Pentru reducerea emisiilor de particule de cenușă în aer, se recomandă montarea "
        "ciclonului, în special dacă beneficiarul a instalat și sistem de curățare pneumatică."
    ),
    (
        "Kod velikih sistema gde se dnevna potrošnja peleta kreće",
        "La sistemele mari, unde consumul zilnic de peleți variază de la câteva sute de "
        "kilograme până la câteva tone, se recomandă montarea unui siloz mare cu elevator "
        "cu cupe. Acesta este conectat la silozul mic prin transportoare melcate, iar întregul "
        "proces de alimentare este automatizat prin sonde de minim și maxim în silozul mic."
    ),
    (
        "Standardni TKAN kotlovi opremljeni su dnevnim silosom",
        "Cazanele standard TKAN sunt echipate cu siloz zilnic, iar pentru sisteme mai mari "
        "producătorul permite realizarea unor silozuri externe de capacitate mai mare, cu "
        "dimensiuni speciale. În funcție de necesarul sistemului, se pot realiza silozuri cu "
        "capacitate de câteva zeci de tone, cu elevator cu cupe."
    ),
    (
        "Veliki spoljni silos povezuje se sa dnevnim silosom",
        "Silozul exterior mare se conectează la silozul zilnic al cazanului prin transportoare "
        "melcate, ceea ce permite alimentarea automată a silozului zilnic prin sonde de minim "
        "și maxim. Pentru depozitarea unor cantități mai mari de peleți se pot utiliza și saci Jumbo."
    ),
    (
        "Kod novijih modela TKAN 60",
        "La modelele mai noi TKAN 60-300, silozul zilnic standard are volumul de 800 litri, "
        "cu posibilitatea conectării la un siloz exterior mare. Conectarea poate fi realizată "
        "lateral sau frontal, în funcție de dispunerea echipamentelor și de condițiile de spațiu. "
        "Schema este:"
    ),
    (
        "Proizvođač predviđa: pužne transportere",
        "Producătorul prevede transportoare melcate, acționări/motoare pentru transportoare, "
        "conectarea silozului mare cu silozul zilnic, completarea automată a silozului zilnic "
        "și sonde de minim și maxim în silozul zilnic."
    ),
    (
        "Kod velikih sistema preporučuje se veliki siloz",
        "La sistemele mari se recomandă un siloz mare cu elevator cu cupe, transportoare melcate "
        "și completare automată a silozului zilnic prin sonde de minim și maxim. Pentru depozitarea "
        "unor cantități mai mari de peleți se pot utiliza și saci Jumbo."
    ),
    (
        "Automatika kotla može da upravlja motorom puža",
        "Automatizarea cazanului poate comanda motorul melcului de alimentare din silozul mare."
    ),
    (
        "Za kaskadu više TKAN kotlova",
        "Pentru o cascadă cu mai multe cazane TKAN, această soluție este deosebit de utilă "
        "deoarece permite proiectarea unui sistem central de distribuție a peleților."
    ),
    (
        "Kod automatskog punjenja dnevnog silosa koriste se sonde",
        "La umplerea automată a silozului zilnic se utilizează sonde de minim și maxim. "
        "Principiul este: MIN -> pornește transportul peleților, MAX -> oprește transportul "
        "peleților. În acest fel cazanul solicită singur combustibil din depozitul central și "
        "nu necesită completare manuală."
    ),
]

RO_SKIP_PREFIXES = (
    "pravni naslednik zanatske radnje",
    "osnovana 1991.",
    "montaža i održavanje",
    "toplovodni kotao",
    "godine.",
    "iz godine u godinu",
    "trudeći da bude",
    "kvalitetu proizvoda",
    "calitateu proizvoda",
    "proizvoda i usluga",
    "proizvodnja povećava",
    "došli do nivoa",
    "svetskim tehnologijama",
    "izdvajaju se",
    "CNC probijanje",
    "kao i upotrebom",
    "perforare CNC.",
    "radnika od kojih",
    "rade na usavršavanju",
    "poslovanja firme",
    "sistema kvaliteta",
    "sistema calitatea",
    "više kotlova, duži dimovod",
    "MIN → uključi transport",
    "MAX → zaustavi transport",
    "Kotao TKAN je razvijen",
)

RO_POST_PREFIX_TRANSLATIONS = [
    (
        "Proizvođač predviđa:",
        "Producătorul prevede transportoare melcate, acționări/motoare pentru transportoare, "
        "conectarea silozului mare cu silozul zilnic, completarea automată a silozului zilnic "
        "și sonde de minim și maxim în silozul zilnic.",
    ),
]

RO_POST_SKIP_PREFIXES = (
    "više kotlova, duži dimovod",
    "Na taj način kotao sam traži",
    "MAX → zaustavi transport",
    "pužne transportere",
    "pogone/motore pužnih transportera",
    "povezivanje velikog i dnevnog",
    "automatsko dopunjavanje dnevnog",
    "automat dopunjavanje dnevnog",
    "sonde minimuma",
)


def normalize_catalog_text(text: str) -> str:
    return (
        text.replace("kotlove snage od 15 do 500 kW", "kotlove snage od 6 do 600 kW")
        .replace("snage od 15 do 500 kW", "snage od 6 do 600 kW")
        .replace("15 do 500 kW", "6 do 600 kW")
        .replace("15-500 kW", "6-600 kW")
        .replace("Dostupni su u osegu od 80 – 500 kW.", "Dostupni su u opsegu snaga od 60 do 300 kW.")
        .replace("Dostupni su u opsegu od 80 – 500 kW.", "Dostupni su u opsegu snaga od 60 do 300 kW.")
        .replace("Dostupni su u opsegu snaga od 80 do 500 kW.", "Dostupni su u opsegu snaga od 60 do 300 kW.")
        .replace("Dostupan je u opsegu snaga od 80 do 500 kW.", "Dostupan je u opsegu snaga od 80 do 600 kW.")
        .replace("(U ponudi snaga od 80 – 500 [kW] )", "(U ponudi snaga od 80 – 600 [kW] )")
        .replace("(U ponudi snaga od 80 - 500 [kW] )", "(U ponudi snaga od 80 - 600 [kW] )")
    )


def translate_text(text: str, language: str) -> str:
    text = normalize_catalog_text(text)
    if language == "sr":
        return text
    if text.startswith(RO_SKIP_PREFIXES):
        return ""
    if text in RO_EXACT_TRANSLATIONS:
        return RO_EXACT_TRANSLATIONS[text]
    for prefix, translation in RO_PREFIX_TRANSLATIONS:
        if text.startswith(prefix):
            return translation
    translated = text
    for source, target in sorted(RO_PHRASE_TRANSLATIONS, key=lambda item: len(item[0]), reverse=True):
        translated = translated.replace(source, target)
    if translated.startswith(RO_POST_SKIP_PREFIXES):
        return ""
    for prefix, translation in RO_POST_PREFIX_TRANSLATIONS:
        if translated.startswith(prefix):
            return translation
    return translated


def render_language_switch(current_language: str) -> str:
    links = []
    for language, label in LANGUAGE_LABELS.items():
        config = LANGUAGE_CONFIG[language]
        classes = "is-active" if language == current_language else ""
        aria_current = ' aria-current="page"' if language == current_language else ""
        href = config["html_path"].name
        links.append(f'<a class="{classes}" href="{href}"{aria_current}>{label}</a>')
    return '<div class="language-switch" aria-label="Language">{} </div>'.format("".join(links))


def render_pdf_link(class_name: str, label: str, pdf_filename: str) -> str:
    if PDF_DOWNLOAD_ENABLED:
        return (
            f'<a class="{class_name}" href="{html.escape(pdf_filename)}">'
            f"{html.escape(label)}</a>"
        )
    return (
        f'<a class="{class_name} is-disabled" aria-disabled="true" tabindex="-1" '
        f'title="PDF je privremeno nedostupan">{html.escape(label)}</a>'
    )


def iter_blocks(document: Document):
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def uppercase_ratio(text: str) -> float:
    letters = [character for character in text if character.isalpha()]
    if not letters:
        return 0
    return sum(1 for character in letters if character.upper() == character) / len(letters)


def is_heading(paragraph: Paragraph, text: str) -> bool:
    if not text or len(text) > 110:
        return False
    if paragraph.style and paragraph.style.name.startswith("List"):
        return False
    return uppercase_ratio(text) > 0.78 or text.lower() == "o nama"


def is_subheading(paragraph: Paragraph, text: str) -> bool:
    if not text or len(text) > 90:
        return False
    if paragraph.style and paragraph.style.name.startswith("List"):
        return uppercase_ratio(text) > 0.78
    prefixes = (
        "serija ",
        "kaskadni ",
        "dodatna ",
        "primena ",
        "specifičnosti ",
        "hidrauličko ",
        "pozicioniranje ",
    )
    return text.lower().startswith(prefixes)


def table_to_html(table: Table, language: str) -> str:
    rows_html: list[str] = []
    for row_index, row in enumerate(table.rows):
        cells = [translate_text(clean_text(cell.text), language) for cell in row.cells]
        tag = "th" if row_index == 0 else "td"
        cell_html = "".join(f"<{tag}>{html.escape(cell)}</{tag}>" for cell in cells)
        rows_html.append(f"<tr>{cell_html}</tr>")
    header = rows_html[0] if rows_html else ""
    body = "".join(rows_html[1:])
    row_count = len(rows_html)
    return (
        f"<div class=\"table-scroll table-scroll--keep table-scroll--rows-{row_count}\" "
        f"data-table-rows=\"{row_count}\"><table>"
        f"<thead>{header}</thead><tbody>{body}</tbody>"
        "</table></div>"
    )


def remove_light_background(image: Image.Image) -> Image.Image:
    rgba = image.convert("RGBA")
    width, height = rgba.size
    data = np.array(rgba)
    rgb = data[:, :, :3].astype(np.int16)
    alpha = data[:, :, 3]
    brightness = rgb.mean(axis=2)
    neutral = rgb.max(axis=2) - rgb.min(axis=2) <= 42
    light_candidate = (brightness >= 238) & neutral & (alpha > 0)

    # Remove only the light background connected to the outside of the image.
    # This preserves highlights, labels and bright details inside the boiler drawings.
    connected = np.zeros((height, width), dtype=bool)
    queue: deque[tuple[int, int]] = deque()

    def enqueue(x: int, y: int) -> None:
        if light_candidate[y, x] and not connected[y, x]:
            connected[y, x] = True
            queue.append((x, y))

    for x in range(width):
        enqueue(x, 0)
        enqueue(x, height - 1)
    for y in range(height):
        enqueue(0, y)
        enqueue(width - 1, y)

    while queue:
        x, y = queue.popleft()
        if x > 0:
            enqueue(x - 1, y)
        if x < width - 1:
            enqueue(x + 1, y)
        if y > 0:
            enqueue(x, y - 1)
        if y < height - 1:
            enqueue(x, y + 1)

    data[connected, 3] = 0
    rgba = Image.fromarray(data, "RGBA")

    bbox = rgba.getbbox()
    if not bbox:
        return rgba
    left, top, right, bottom = bbox
    padding = max(18, int(min(width, height) * 0.025))
    crop_box = (
        max(0, left - padding),
        max(0, top - padding),
        min(width, right + padding),
        min(height, bottom + padding),
    )
    return rgba.crop(crop_box)


def save_display_image(raw_path: Path, display_path: Path) -> None:
    with Image.open(raw_path) as image:
        image.seek(0)
        cleaned = remove_light_background(image)
        max_side = max(cleaned.size)
        if max_side < 1400:
            scale = min(2.0, 1400 / max_side)
            cleaned = cleaned.resize(
                (round(cleaned.width * scale), round(cleaned.height * scale)),
                Image.Resampling.LANCZOS,
            ).filter(ImageFilter.UnsharpMask(radius=1.1, percent=115, threshold=3))
        cleaned.save(display_path, optimize=False, compress_level=4)


def prepare_supplemental_images() -> None:
    EDITORIAL_ASSET_DIR.mkdir(parents=True, exist_ok=True)
    for target_name, source_path in SUPPLEMENTAL_IMAGE_SOURCES.items():
        if not source_path.is_file():
            raise FileNotFoundError(f"Nedostaje slika za katalog: {source_path}")
        save_display_image(source_path, EDITORIAL_ASSET_DIR / target_name)


def extract_images() -> list[dict[str, str]]:
    if ASSET_DIR.exists():
        shutil.rmtree(ASSET_DIR)
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    images: list[dict[str, str]] = []

    with zipfile.ZipFile(DOCX) as archive:
        media_entries = [
            entry for entry in archive.infolist() if entry.filename.startswith("word/media/")
        ]
        for index, entry in enumerate(media_entries, start=1):
            original_name = Path(entry.filename).name
            suffix = Path(original_name).suffix.lower()
            raw_path = ASSET_DIR / original_name
            raw_path.write_bytes(archive.read(entry))

            target_name = f"catalog-image-{index:02d}"
            display_path: Path | None = None
            status = "included"

            if suffix in {".png", ".jpg", ".jpeg"}:
                display_path = ASSET_DIR / f"{target_name}.png"
                save_display_image(raw_path, display_path)
            elif suffix in {".tif", ".tiff"}:
                display_path = ASSET_DIR / f"{target_name}.png"
                save_display_image(raw_path, display_path)
            elif suffix == ".wmf":
                # Browser support for WMF is not reliable; keep the original asset visible as a download item.
                status = "original-wmf"
                display_path = None
            else:
                status = "unsupported"

            images.append(
                {
                    "label": f"Slika {index}",
                    "original": original_name,
                    "key": Path(original_name).name.lower(),
                    "display": display_path.name if display_path else "",
                    "status": status,
                }
            )

    return images


def paragraph_image_keys(paragraph: Paragraph) -> list[str]:
    keys: list[str] = []
    for blip in paragraph._element.xpath('.//*[local-name()="blip"]'):
        rel_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if not rel_id:
            continue
        rel = paragraph.part.rels.get(rel_id)
        if rel is not None:
            keys.append(Path(rel.target_ref).name.lower())
    return keys


def has_word_page_break(paragraph: Paragraph) -> bool:
    return bool(
        paragraph._element.xpath('.//*[local-name()="lastRenderedPageBreak"]')
        or paragraph._element.xpath(
            './/*[local-name()="br" and @*[local-name()="type"]="page"]'
        )
    )


def image_figure(item: dict[str, str], index: int, language: str) -> str:
    label = "Slika" if language == "sr" else "Imagine"
    if not item.get("display"):
        return (
            "<div class=\"catalog-original-note\">"
            f"<strong>{label} {index}</strong>: originalni fajl "
            f"{html.escape(item['original'])} je sacuvan u assets/full-catalog."
            "</div>"
        )
    return f"""
    <figure class="catalog-figure">
      <img src="assets/full-catalog/{html.escape(item['display'])}" alt="{label} {index}" />
    </figure>
    """


def render_copy(blocks: list[dict[str, object]]) -> str:
    rendered: list[str] = []
    for block in blocks:
        block_type = block["type"]
        break_class = " class=\"word-page-break\"" if block.get("page_break_before") else ""
        if block_type == "paragraph":
            rendered.append(f"<p{break_class}>{html.escape(str(block['text']))}</p>")
        elif block_type == "subheading":
            rendered.append(f"<h3{break_class}>{html.escape(str(block['text']))}</h3>")
        elif block_type == "list":
            items = "".join(f"<li>{html.escape(str(item))}</li>" for item in block["items"])
            classes = "feature-list word-page-break" if block.get("page_break_before") else "feature-list"
            rendered.append(f"<ul class=\"{classes}\">{items}</ul>")
    return "".join(rendered)


def render_section(section: dict[str, object], section_number: int, language: str) -> str:
    blocks = list(section["blocks"])
    rendered: list[str] = []
    media_index = 0
    index = 0

    while index < len(blocks):
        block = blocks[index]
        block_type = block["type"]

        if block_type in {"paragraph", "subheading", "list"}:
            copy_end = index
            while copy_end < len(blocks) and blocks[copy_end]["type"] in {"paragraph", "subheading", "list"}:
                copy_end += 1
            if copy_end < len(blocks) and blocks[copy_end]["type"] == "figure":
                media_index += 1
                side = "media-block--image-right" if media_index % 2 else "media-block--image-left"
                paired_copy = blocks[index:copy_end]
                page_break_class = (
                    " word-page-break" if paired_copy and paired_copy[0].get("page_break_before") else ""
                )
                after_figure = copy_end + 1
                copy_length = sum(len(str(item.get("text", ""))) for item in paired_copy)
                if copy_length < 260:
                    while (
                        after_figure < len(blocks)
                        and blocks[after_figure]["type"] in {"paragraph", "subheading", "list"}
                        and len(paired_copy) < 3
                    ):
                        paired_copy.append(blocks[after_figure])
                        after_figure += 1
                rendered.append(
                    f"<div class=\"media-block {side}{page_break_class}\">"
                    f"<div class=\"media-copy\">{render_copy(paired_copy)}</div>"
                    f"{blocks[copy_end]['html']}"
                    "</div>"
                )
                index = after_figure
                continue
            rendered.append(render_copy(blocks[index:copy_end]))
            index = copy_end
            continue

        if block_type == "figure":
            copy_end = index + 1
            while copy_end < len(blocks) and blocks[copy_end]["type"] in {"paragraph", "subheading", "list"}:
                copy_end += 1
            if copy_end > index + 1:
                media_index += 1
                side = "media-block--image-left" if media_index % 2 else "media-block--image-right"
                page_break_class = " word-page-break" if block.get("page_break_before") else ""
                rendered.append(
                    f"<div class=\"media-block {side}{page_break_class}\">"
                    f"{block['html']}"
                    f"<div class=\"media-copy\">{render_copy(blocks[index + 1:copy_end])}</div>"
                    "</div>"
                )
                index = copy_end
                continue
            page_break_class = " word-page-break" if block.get("page_break_before") else ""
            rendered.append(
                f"<div class=\"technical-visual{page_break_class}\">{block['html']}</div>"
            )
        elif block_type == "table":
            rendered.append(str(block["html"]))
        index += 1

    section_body = "".join(rendered)
    if section_number == 1:
        gallery_label = (
            "Radijator Inženjering producție"
            if language == "ro"
            else "Radijator Inženjering proizvodnja"
        )
        section_body = (
            '<div class="about-layout">'
            f'<div class="about-copy">{section_body}</div>'
            f'<aside class="about-gallery" aria-label="{gallery_label}">'
            '<figure class="about-photo about-photo--company">'
            '<img src="assets/editorial/about-factory-exterior.jpg" '
            'alt="Radijator Inženjering factory complex in Kraljevo" />'
            '</figure>'
            '<figure class="about-photo">'
            '<img src="assets/editorial/about-production-forming.jpg" '
            'alt="Automated sheet metal processing in production" />'
            '</figure>'
            '<figure class="about-photo">'
            '<img src="assets/editorial/about-production-laser.jpg" '
            'alt="Laser cutting of boiler sheet metal" />'
            '</figure>'
            '<figure class="about-photo">'
            '<img src="assets/editorial/about-production-control.jpg" '
            'alt="Operator supervising a modern production process" />'
            '</figure>'
            '<figure class="about-photo">'
            '<img src="assets/editorial/about-production-welding.jpg" '
            'alt="Welding industrial boiler components" />'
            '</figure>'
            '</aside>'
            '</div>'
        )

    media_class = " catalog-section--media" if any(block["type"] == "figure" for block in blocks) else ""
    return (
        f"<section class=\"catalog-section{media_class}\" id=\"{section['id']}\" data-section=\"{section_number:02d}\">"
        f"<div class=\"section-heading\"><span>{section_number:02d}</span>"
        f"<h2>{html.escape(str(section['title']))}</h2></div>"
        f"{section_body}"
        "</section>"
    )


def render_production_spread(language: str) -> str:
    if language == "ro":
        return """
<section class="catalog-section production-spread" id="production-standards" data-section="PRO">
  <div class="production-spread__brand">
    <span>Producție</span>
  </div>
  <div class="production-spread__headline">
    <p class="production-spread__kicker">Tehnologie și calitate</p>
    <h2>Producție conform standardelor europene moderne</h2>
  </div>
  <div class="production-spread__media">
    <figure><img src="assets/editorial/company-aerial-complex-wide.jpg" alt="Complexul de producție Radijator Inzenjering văzut din aer" /></figure>
    <figure><img src="assets/editorial/company-aerial-complex-top.jpg" alt="Fabrica Radijator Inzenjering cu producție modernă" /></figure>
  </div>
  <div class="production-spread__copy">
    <p>Pe măsură ce producția s-a extins și s-a perfecționat, cazanele au început să fie fabricate cu cele mai moderne tehnologii: tăiere laser, procedeu CNC plasmă, perforare CNC, sudură robotică și sudură automatizată.</p>
    <p>Astăzi Radijator Inženjering are peste 350 de angajați, printre care 40 de ingineri mecanici absolvenți care lucrează zilnic la îmbunătățirea calității produselor.</p>
  </div>
  <div class="production-spread__stats">
    <div><strong>350+</strong><span>angajați</span></div>
    <div><strong>40</strong><span>ingineri mecanici absolvenți</span></div>
    <div><strong>EU</strong><span>export în 27+ țări UE</span></div>
  </div>
  <p class="production-spread__footer">Tehnologie / calitate / piață</p>
</section>
"""
    return """
<section class="catalog-section production-spread" id="production-standards" data-section="PRO">
  <div class="production-spread__brand">
    <span>Proizvodnja</span>
  </div>
  <div class="production-spread__headline">
    <p class="production-spread__kicker">Tehnologija i kvalitet</p>
    <h2>Proizvodnja po savremenim evropskim standardima</h2>
  </div>
  <div class="production-spread__media">
    <figure><img src="assets/editorial/company-aerial-complex-wide.jpg" alt="Proizvodni kompleks Radijator Inzenjering iz vazduha" /></figure>
    <figure><img src="assets/editorial/company-aerial-complex-top.jpg" alt="Pogon Radijator Inzenjering sa savremenom proizvodnjom" /></figure>
  </div>
  <div class="production-spread__copy">
    <p>Kako se proizvodnja sirila i usavrsavala, kotlovi su poceli da se izradjuju najsavremenijim tehnologijama: lasersko secenje, CNC plazma postupak, CNC probijanje, robotsko zavarivanje i zavarivanje automatima.</p>
    <p>Danas Radijator Inzenjering zaposljava preko 350 radnika, medju kojima je 40 diplomiranih masinskih inzenjera koji svakodnevno rade na unapredjenju kvaliteta proizvoda.</p>
  </div>
  <div class="production-spread__stats">
    <div><strong>350+</strong><span>zaposlenih</span></div>
    <div><strong>40</strong><span>dipl. masinskih inzenjera</span></div>
    <div><strong>EU</strong><span>izvoz u 27+ zemalja EU</span></div>
  </div>
  <p class="production-spread__footer">Tehnologija / kvalitet / trziste</p>
</section>
"""


def render_boiler_room_figure() -> str:
    return (
        '<div class="technical-visual boiler-room-visual">'
        '<div class="figure-row">'
        '<figure class="catalog-figure">'
        '<img src="assets/editorial/boiler-room-position.png" '
        'alt="Pozicioniranje TKAN kotla u kotlarnici" />'
        '</figure>'
        '</div>'
        '</div>'
    )


def render_supplemental_visuals(items: list[tuple[str, str]], modifier: str) -> str:
    figures = "\n".join(
        (
            '<figure class="supplemental-visuals__item">'
            f'<img src="assets/editorial/{html.escape(filename)}" alt="{html.escape(alt)}" />'
            "</figure>"
        )
        for filename, alt in items
    )
    return (
        f'<div class="supplemental-visuals supplemental-visuals--{modifier}">'
        f"{figures}"
        "</div>"
    )


def append_after_section_table(body_html: str, section_id: str, insert_html: str) -> str:
    section_start = body_html.find(f'id="{section_id}"')
    section_end = body_html.find("</section>", section_start)
    if section_start == -1 or section_end == -1:
        return body_html
    section_html = body_html[section_start:section_end]
    table_end = section_html.find("</table></div>")
    if table_end == -1:
        return body_html
    table_end += len("</table></div>")
    updated_section = section_html[:table_end] + insert_html + section_html[table_end:]
    return body_html[:section_start] + updated_section + body_html[section_end:]


def tune_catalog_layout(body_html: str) -> str:
    """Apply editorial moves that keep generated content aligned with the catalog story."""
    section_start = body_html.find('id="section-07"')
    section_end = body_html.find("</section>", section_start)
    if section_start != -1 and section_end != -1:
        section_html = body_html[section_start:section_end]
        table_start = section_html.find('<div class="table-scroll table-scroll--keep table-scroll--rows-13"')
        trailing_media_start = section_html.find(
            '<div class="media-block media-block--image-left">',
            table_start,
        )
        if table_start != -1 and trailing_media_start != -1:
            table_html = section_html[table_start:trailing_media_start]
            trailing_media_html = section_html[trailing_media_start:]
            trailing_media_html = trailing_media_html.replace("<p>.</p>", "")
            reordered_section = (
                section_html[:table_start]
                + trailing_media_html
                + table_html
            )
            body_html = (
                body_html[:section_start]
                + reordered_section
                + body_html[section_end:]
            )

    section_start = body_html.find('id="section-10"')
    section_end = body_html.find("</section>", section_start)
    if section_start != -1 and section_end != -1:
        section_html = body_html[section_start:section_end]
        heading_end = section_html.find("</div>")
        first_p_start = section_html.find("<p>", heading_end)
        first_p_end = section_html.find("</p>", first_p_start)
        table_start = section_html.find('<div class="table-scroll', first_p_end)
        if heading_end != -1 and first_p_start != -1 and first_p_end != -1 and table_start != -1:
            intro_html = section_html[first_p_start:first_p_end + 4]
            section_without_intro = (
                section_html[:first_p_start]
                + section_html[first_p_end + 4:]
            )
            table_end = section_without_intro.find("</table></div>", table_start - len(intro_html))
            if table_end != -1 and "boiler-room-visual" not in section_without_intro:
                table_end += len("</table></div>")
                figure_html = render_boiler_room_figure()
                section_without_intro = (
                    section_without_intro[:heading_end + 6]
                    + figure_html
                    + section_without_intro[heading_end + 6:table_end]
                    + intro_html
                    + section_without_intro[table_end:]
                )
                body_html = (
                    body_html[:section_start]
                    + section_without_intro
                    + body_html[section_end:]
                )

    body_html = body_html.replace(
        '    </figure>\n'
        '    </div></div><div class="technical-visual"><div class="figure-row">\n'
        '    <figure class="catalog-figure">\n'
        '      <img src="assets/full-catalog/catalog-image-14.png" alt="Slika 14" />\n'
        '    </figure>\n'
        '    </div></div></section>',
        '    </figure>\n'
        '    <figure class="catalog-figure">\n'
        '      <img src="assets/full-catalog/catalog-image-14.png" alt="Slika 14" />\n'
        '    </figure>\n'
        '    </div></div></section>',
    )
    body_html = body_html.replace(
        '    </figure>\n'
        '    </div></div><div class="technical-visual"><div class="figure-row">\n'
        '    <figure class="catalog-figure">\n'
        '      <img src="assets/full-catalog/catalog-image-14.png" alt="Imagine 19" />\n'
        '    </figure>\n'
        '    </div></div></section>',
        '    </figure>\n'
        '    <figure class="catalog-figure">\n'
        '      <img src="assets/full-catalog/catalog-image-14.png" alt="Imagine 19" />\n'
        '    </figure>\n'
        '    </div></div></section>',
    )

    body_html = append_after_section_table(
        body_html,
        "section-03",
        render_supplemental_visuals(
            [
                ("valvola-tkan-150-presek.png", "Valvola TKAN 150 - presek"),
                ("valvola-tkan-150.png", "Valvola TKAN 150"),
            ],
            "two",
        ),
    )
    body_html = append_after_section_table(
        body_html,
        "section-07",
        render_supplemental_visuals(
            [
                ("valvola-tkan-300-integra-presek.png", "Valvola TKAN 300 Integra - presek"),
                ("valvola-tkan-300-integra.png", "Valvola TKAN 300 Integra"),
                ("multiciklon-tkan-300.png", "Multiciklon TKAN 300"),
            ],
            "three",
        ),
    )

    return body_html


def tune_romanian_catalog_html(body_html: str) -> str:
    """Clean Word fragments that survive paragraph-level translation in Romanian output."""
    replacements = {
        "<h3>Kaskadni sistemi</h3>": "<h3>Sisteme în cascadă</h3>",
        (
            "Kaskadni sistemi predstavljaju kombinaciju dva ili više kotlova povezanih u "
            "jedinstven sistem sa zajedničkim silozom za skladištenje peleți. Ovakva "
            "konfiguracija omogućava veću ukupnu instalisanu snagu, pouzdaniji rad, "
            "ravnomernu raspodelu opterećenja i veću energetsku eficiență sistema."
        ): (
            "Sistemele în cascadă reprezintă o combinație de două sau mai multe cazane "
            "conectate într-un sistem unic, cu siloz comun pentru depozitarea peleților. "
            "O astfel de configurație permite o putere totală instalată mai mare, "
            "funcționare mai fiabilă, distribuție uniformă a sarcinii și eficiență "
            "energetică mai ridicată a sistemului."
        ),
        "<h3>Dodatna oprema</h3>": "<h3>Echipamente suplimentare</h3>",
        (
            "Pored kotlova, u ponudi je kompletna prateća oprema za formiranje "
            "funkcionalnog sistema grejanja. Asortiman obuhvata siloze za skladištenje "
            "peleți, pužne transportere, elevatore, bafer rezervoare, automatiku i ostalu "
            "opremu potrebnu za pouzdan transport, doziranje i kontrolu peleți, kao i "
            "siguran i efikasan rad celokupnog sistema."
        ): (
            "Pe lângă cazane, oferta include echipamente auxiliare complete pentru formarea "
            "unui sistem funcțional de încălzire. Gama cuprinde silozuri pentru depozitarea "
            "peleților, transportoare melcate, elevatoare, rezervoare tampon, automatizare "
            "și alte echipamente necesare pentru transportul, dozarea și controlul fiabil "
            "al peleților, precum și pentru funcționarea sigură și eficientă a întregului sistem."
        ),
        (
            "<p>više kotlova, duži dimovod, ciklon/multiciklon, veći broj kolena, "
            "zajednički dimnjak.</p>"
        ): "",
        (
            "<p>Proizvođač predviđa: pužne transportere, pogone/motore pužnih transportera, "
            "povezivanje velikog i dnevnog silozului, automat dopunjavanje dnevnog silozului, "
            "sonde minimuma i maksimuma u dnevnom silozu.</p>"
        ): (
            "<p>Producătorul prevede transportoare melcate, acționări/motoare pentru "
            "transportoare, conectarea silozului mare cu silozul zilnic, completarea "
            "automată a silozului zilnic și sonde de minim și maxim în silozul zilnic.</p>"
        ),
        (
            "<p>Na taj način kotao sam traži gorivo iz centralnog skladišta i ne zahteva "
            "ručno dopunjavanje.</p>"
        ): "",
    }
    for source, target in replacements.items():
        body_html = body_html.replace(source, target)
    return body_html


def build_content(images_by_key: dict[str, dict[str, str]], language: str) -> tuple[str, list[str], int, int, set[str]]:
    document = Document(DOCX)
    sections: list[dict[str, object]] = []
    toc: list[str] = []
    used_images: set[str] = set()
    current_section: dict[str, object] | None = None
    paragraph_count = 0
    table_count = 0
    figure_count = 0
    paragraph_buffer: list[str] = []
    list_buffer: list[str] = []
    paragraph_break_before = False
    list_break_before = False

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer, paragraph_break_before
        if paragraph_buffer and current_section is not None:
            current_section["blocks"].append(
                {
                    "type": "paragraph",
                    "text": " ".join(paragraph_buffer),
                    "page_break_before": paragraph_break_before,
                }
            )
            paragraph_buffer = []
            paragraph_break_before = False

    def flush_list() -> None:
        nonlocal list_buffer, list_break_before
        if list_buffer and current_section is not None:
            current_section["blocks"].append(
                {
                    "type": "list",
                    "items": list_buffer,
                    "page_break_before": list_break_before,
                }
            )
            list_buffer = []
            list_break_before = False

    def ensure_section(title: str | None = None) -> dict[str, object]:
        nonlocal current_section
        if current_section is None:
            section_title = title or translate_text("O nama", language)
            section_id = f"section-{len(sections) + 1:02d}"
            current_section = {"id": section_id, "title": section_title, "blocks": []}
            sections.append(current_section)
            toc.append(f"<a href=\"#{section_id}\">{html.escape(section_title)}</a>")
        return current_section

    def start_section(title: str) -> None:
        nonlocal current_section
        flush_paragraph()
        flush_list()
        section_title = translate_text(title, language)
        section_id = f"section-{len(sections) + 1:02d}"
        current_section = {"id": section_id, "title": section_title, "blocks": []}
        sections.append(current_section)
        toc.append(f"<a href=\"#{section_id}\">{html.escape(section_title)}</a>")

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            source_text = clean_text(block.text)
            text = translate_text(source_text, language)
            image_keys = paragraph_image_keys(block)
            page_break_before = has_word_page_break(block)
            if not source_text and not image_keys:
                continue
            if is_heading(block, source_text):
                start_section(source_text)
            else:
                section = ensure_section()
                if text:
                    paragraph_count += 1
                    if is_subheading(block, source_text):
                        flush_paragraph()
                        flush_list()
                        section["blocks"].append(
                            {
                                "type": "subheading",
                                "text": text,
                                "page_break_before": page_break_before,
                            }
                        )
                    elif block.style and block.style.name.startswith("List"):
                        flush_paragraph()
                        if page_break_before:
                            flush_list()
                            list_break_before = True
                        list_buffer.append(text)
                    else:
                        flush_list()
                        if page_break_before:
                            flush_paragraph()
                            paragraph_break_before = True
                        paragraph_buffer.append(text)
                        ends_with_year = bool(re.search(r"\b(?:19|20)\d{2}\.\s*$", text))
                        if (re.search(r"[.!?]\s*$", text) and not ends_with_year) or len(text) > 240:
                            flush_paragraph()
                if image_keys:
                    flush_paragraph()
                    flush_list()
                    figure_group: list[str] = []
                    for key in image_keys:
                        item = images_by_key.get(key)
                        if not item:
                            continue
                        used_images.add(key)
                        figure_count += 1
                        figure_group.append(image_figure(item, figure_count, language))
                    if figure_group:
                        section["blocks"].append(
                            {"type": "figure", "html": f"<div class=\"figure-row\">{''.join(figure_group)}</div>"}
                        )
                        section["blocks"][-1]["page_break_before"] = page_break_before
        elif isinstance(block, Table):
            section = ensure_section()
            flush_paragraph()
            flush_list()
            table_count += 1
            section["blocks"].append({"type": "table", "html": table_to_html(block, language)})

    flush_paragraph()
    flush_list()
    rendered_sections = [
        render_section(section, index, language)
        for index, section in enumerate(sections, start=1)
    ]
    return "\n".join(rendered_sections), toc, paragraph_count, table_count, used_images


def render_page(language: str) -> None:
    config = LANGUAGE_CONFIG[language]
    prepare_supplemental_images()
    images = extract_images()
    images_by_key = {item["key"]: item for item in images}
    body_html, toc, _, _, _ = build_content(images_by_key, language)
    body_html = tune_catalog_layout(body_html)
    if language == "ro":
        body_html = tune_romanian_catalog_html(body_html)
    production_spread = render_production_spread(language)
    body_html = body_html.replace("</section>", f"</section>\n{production_spread}", 1)
    toc.insert(1, f'<a href="#production-standards">{html.escape(config["production_toc"])}</a>')
    language_switch = render_language_switch(language)
    nav_pdf_link = render_pdf_link("web-nav-pdf", config["pdf_label"], config["pdf_filename"])
    hero_pdf_link = render_pdf_link("action-secondary", config["hero_secondary"], config["pdf_filename"])
    print_title_override = ""
    if language == "ro":
        print_title_override = """
    <style>
      @media print {
        @page {
          @top-center { content: "CAZANE INDUSTRIALE PE BIOMASĂ"; }
        }
        @page catalog {
          @top-center { content: "CAZANE INDUSTRIALE PE BIOMASĂ"; }
        }
      }
    </style>"""

    page = f"""<!doctype html>
<html lang="{config["html_lang"]}">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{html.escape(config["page_title"])}</title>
    <link rel="icon" href="assets/favicon.svg" />
    <link rel="stylesheet" href="styles.css" />
    <link rel="stylesheet" href="catalog-premium.css" />
    <link rel="stylesheet" href="catalog-web.css" />
{print_title_override}
  </head>
  <body class="catalog-page" id="top">
    <div class="catalog-progress" aria-hidden="true"><span></span></div>
    <nav class="catalog-web-nav" aria-label="{html.escape(config["aria_main"])}">
      <a class="web-nav-brand" href="#top" aria-label="{html.escape(config["brand_aria"])}">
        <img src="assets/logo.png" alt="" />
        <span>{html.escape(config["brand_title"])}</span>
      </a>
      <div class="web-nav-links">
        <a href="#section-03">{html.escape(config["nav_boilers"])}</a>
        <a href="#section-11">{html.escape(config["nav_systems"])}</a>
        <a href="#section-12">{html.escape(config["nav_equipment"])}</a>
        <a href="#kontakt">{html.escape(config["nav_contact"])}</a>
      </div>
      {language_switch}
      {nav_pdf_link}
    </nav>
    <header class="catalog-hero">
      <div class="catalog-hero-topline">
        <div class="catalog-logo-card">
          <img src="assets/logo.png" alt="Radijator Inzenjering" />
        </div>
        <span>{html.escape(config["hero_topline"])}</span>
      </div>
      <div class="catalog-hero-copy">
        <p class="eyebrow">{html.escape(config["hero_eyebrow"])}</p>
        <h1>{config["hero_title"]}</h1>
        <p class="catalog-lead">{html.escape(config["hero_lead"])}</p>
        <div class="catalog-actions">
          <a class="action-primary" href="#section-01">{html.escape(config["hero_primary"])}</a>
          {hero_pdf_link}
        </div>
      </div>
      <div class="hero-machine">
        <span class="machine-orbit machine-orbit--outer"></span>
        <span class="machine-orbit machine-orbit--inner"></span>
        <div class="hero-live-frame">
          <figure class="hero-anniversary-card">
            <img src="assets/editorial/anniversary-badge.png" alt="{html.escape(config["badge_alt"])}" fetchpriority="high" />
          </figure>
        </div>
      </div>
      <a class="scroll-cue" href="#section-01" aria-label="{html.escape(config["hero_primary"])}"><span></span>{html.escape(config["scroll"])}</a>
    </header>
    <main class="catalog-layout">
      <details class="catalog-toc" open>
        <summary>{html.escape(config["toc_summary"])}</summary>
        <a class="back-link" href="#top">{html.escape(config["top_link"])}</a>
        <nav>{"".join(toc)}</nav>
      </details>
      <article class="catalog-content">
        {body_html}
      </article>
    </main>
    <footer class="catalog-footer" id="kontakt">
      <div class="catalog-footer-main">
        <p class="footer-kicker">{html.escape(config["footer_kicker"])}</p>
        <h2>{html.escape(config["footer_title"])}</h2>
        <a class="footer-mail" href="mailto:radijator@radijator.rs">radijator@radijator.rs</a>
      </div>
      <div class="catalog-footer-contact">
        <p><strong>Radijator Inženjering d.o.o.</strong><br />{config["footer_address"]}</p>
        <p><a href="tel:+38136399140">+381 36 399 140</a><br /><a href="https://www.radijator.rs/">www.radijator.rs</a></p>
      </div>
      <section class="catalog-footer-gallery" aria-label="{html.escape(config["footer_gallery"])}">
        <figure class="footer-photo footer-photo--wide"><img src="assets/editorial/hero-boiler-installation.jpg" alt="Instalirani industrijski kotao Radijator u kotlarnici" /></figure>
        <figure class="footer-photo footer-photo--tall"><img src="assets/editorial/hero-boiler-room.jpg" alt="Kaskadno postrojenje sa industrijskim kotlovima Radijator" /></figure>
      </section>
      <div class="catalog-footer-bottom">
        <span>{html.escape(config["footer_product"])}</span>
        <a href="#top">{html.escape(config["footer_top"])}</a>
      </div>
    </footer>
    <dialog class="catalog-lightbox" aria-label="{html.escape(config["dialog_label"])}">
      <button class="lightbox-close" type="button" aria-label="{html.escape(config["dialog_close"])}">{html.escape(config["dialog_close_text"])}</button>
      <div class="lightbox-stage">
        <img alt="" />
        <p></p>
      </div>
    </dialog>
    <script src="catalog.js" defer></script>
  </body>
</html>
"""
    config["html_path"].write_text(page, encoding="utf-8")
    config["alias_path"].write_text(page, encoding="utf-8")


if __name__ == "__main__":
    for selected_language in LANGUAGE_CONFIG:
        render_page(selected_language)
