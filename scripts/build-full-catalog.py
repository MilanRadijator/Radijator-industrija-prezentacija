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
BG_HTML = DOCS / "index-bg.html"
BG_ALIAS_HTML = DOCS / "full-catalog-bg.html"
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
EDITORIAL_PRODUCT_IMAGE_SOURCES = {
    "tkan-300-silos.png": Path.home() / "Downloads" / "TKAN 300 + SILOS.jpg",
    "tkan-integra-render.png": Path.home() / "Downloads" / "TKAN INTEGRA.jpg",
    "kaskadni-sistem-render.png": Path.home() / "Downloads" / "KASKADNI SISTEM.jpg",
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
    "bg": {
        "html_lang": "bg",
        "html_path": BG_HTML,
        "alias_path": BG_ALIAS_HTML,
        "pdf_filename": "radijator-industrijski-kotlovi-bg.pdf",
        "page_title": "Пълен каталог | Radijator Inzenjering",
        "aria_main": "Основна навигация",
        "brand_title": "Индустриален каталог",
        "brand_aria": "Radijator Inženjering - начало на каталога",
        "nav_boilers": "Котли",
        "nav_systems": "Системи",
        "nav_equipment": "Оборудване",
        "nav_contact": "Контакт",
        "pdf_label": "PDF каталог",
        "hero_topline": "Индустриални топлоенергийни решения",
        "hero_eyebrow": "Пълен продуктов каталог / 2026",
        "hero_title": "Индустриални котли <em>на биомаса</em>",
        "hero_lead": (
            "Надеждни системи с висока мощност, проектирани за ефективност, "
            "дълъг експлоатационен живот и пълен контрол на горивния процес."
        ),
        "hero_primary": "Разгледай каталога",
        "hero_secondary": "Изтегли PDF",
        "scroll": "Превъртете",
        "toc_summary": "Съдържание на каталога",
        "top_link": "Начало на каталога",
        "production_toc": "Производство и стандарти",
        "footer_kicker": "Проектиране / производство / поддръжка",
        "footer_title": "Партньор за цялостни топлоенергийни решения.",
        "footer_address": "Živojina Lazića Solunca 6<br />36000 Кралево, Сърбия",
        "footer_gallery": "Radijator Inženjering на практика",
        "footer_product": "Индустриални котли на биомаса",
        "footer_top": "Назад към началото",
        "dialog_label": "Увеличен технически изглед",
        "dialog_close": "Затвори увеличения изглед",
        "dialog_close_text": "Затвори",
        "badge_alt": "35 години опит - качество без компромис",
        "badge_title": "35 години опит",
    },
}

LANGUAGE_LABELS = {"sr": "SR", "ro": "RO", "bg": "BG"}

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
    "Serija TKAN modeli": "Seria TKAN",
    "Serija TKAN": "Seria TKAN",
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
    "Serija TKAN Integra modeli": "Seria TKAN Integra",
    "Serija TKAN Integra": "Seria TKAN Integra",
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
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ TKAN"
    ),
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM TKAN": (
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ TKAN"
    ),
    "PRESEK TKAN KOTLA SA OPISOM ELEMENATA": (
        "SECȚIUNE CAZAN TKAN CU DESCRIEREA ELEMENTELOR"
    ),
    "PRESEK KOTLA TKAN": "SECȚIUNE CAZAN TKAN",
    "TABELA SA DIMENZIJAMA TKAN KOTLA": "TABEL CU DIMENSIUNILE CAZANULUI TKAN",
    "TABELA SA DIMENZIJAMA TKAN SILOSA": "TABEL CU DIMENSIUNILE SILOZULUI TKAN",
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA MODEL": (
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ, DESPRĂFUIRE ȘI CICLON TKAN INTEGRA"
    ),
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA": (
        "CAZAN DE APĂ CALDĂ PE PELEȚI CU ALIMENTARE AUTOMATĂ, DESPRĂFUIRE ȘI CICLON TKAN INTEGRA"
    ),
    "PRESEK TKAN INTEGRA KOTLA SA OPISOM ELEMENATA": (
        "SECȚIUNE CAZAN TKAN INTEGRA CU DESCRIEREA ELEMENTELOR"
    ),
    "PRESEK KOTLA TKAN INTEGRA": "SECȚIUNE CAZAN TKAN INTEGRA",
    "TABELA SA DIMENZIJAMA TKAN INTEGRA SILOSA": (
        "TABEL CU DIMENSIUNILE SILOZULUI TKAN INTEGRA"
    ),
    "POLOŽAJ TKAN OBIČNOG I TKAN INTEGRA KOTLA U KOTLARNICI": (
        "POZIȚIONAREA CAZANELOR TKAN ȘI TKAN INTEGRA ÎN CAMERA TEHNICĂ"
    ),
    "POLOŽAJ TKAN OBIČNOG i TKAN INTEGRA KOTLA U KOTLARNICI": (
        "POZIȚIONAREA CAZANELOR TKAN ȘI TKAN INTEGRA ÎN CAMERA TEHNICĂ"
    ),
    "POLOŽAJ KOTLOVA TKAN i TKAN INTEGRA U KOTLARNICI": (
        "POZIȚIONAREA CAZANELOR TKAN ȘI TKAN INTEGRA ÎN CAMERA TEHNICĂ"
    ),
    "KASKADNI SISTEMI": "SISTEME ÎN CASCADĂ",
    "DODATNA OPREMA": "ECHIPAMENTE SUPLIMENTARE",
    "AUTOMATSKI TRANSPORT PELETA": "TRANSPORT AUTOMAT AL PELEȚILOR",
    "SERIJA TKAN MODELI": "SERIA TKAN",
    "SERIJA TKAN": "SERIA TKAN",
    "SERIJA TKAN INTEGRA MODELI": "SERIA TKAN INTEGRA",
    "SERIJA TKAN INTEGRA": "SERIA TKAN INTEGRA",
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

BG_EXACT_TRANSLATIONS = {
    "O nama": "За нас",
    "PROIZVODNI PROGRAM – INDUSTRIJA": "ПРОИЗВОДСТВЕНА ПРОГРАМА - ИНДУСТРИЯ",
    "PROIZVODNI PROGRAM - INDUSTRIJA": "ПРОИЗВОДСТВЕНА ПРОГРАМА - ИНДУСТРИЯ",
    "Serija TKAN": "Серия TKAN",
    "Serija TKAN Integra": "Серия TKAN Integra",
    "Kaskadni sistemi": "Каскадни системи",
    "Dodatna oprema": "Допълнително оборудване",
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM TKAN": "ВОДОГРЕЕН КОТЕЛ НА ПЕЛЕТИ С АВТОМАТИЧНО ПОДАВАНЕ TKAN",
    "(U ponudi snaga od 60 – 300 [kW] )": "(Предлага се в мощностен диапазон 60 - 300 [kW])",
    "(U ponudi snaga od 60 - 300 [kW] )": "(Предлага се в мощностен диапазон 60 - 300 [kW])",
    "PRESEK KOTLA TKAN": "РАЗРЕЗ НА КОТЕЛ TKAN",
    "TABELA SA DIMENZIJAMA TKAN KOTLA": "ТАБЛИЦА С РАЗМЕРИ НА КОТЕЛ TKAN",
    "TABELA SA DIMENZIJAMA TKAN SILOSA": "ТАБЛИЦА С РАЗМЕРИ НА СИЛОЗ TKAN",
    "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA": "ВОДОГРЕЕН КОТЕЛ НА ПЕЛЕТИ С АВТОМАТИЧНО ПОДАВАНЕ, ОБЕЗПРАШАВАНЕ И ЦИКЛОН - TKAN INTEGRA",
    "(U ponudi snaga od 80 – 600 [kW] )": "(Предлага се в мощностен диапазон 80 - 600 [kW])",
    "(U ponudi snaga od 80 - 600 [kW] )": "(Предлага се в мощностен диапазон 80 - 600 [kW])",
    "PRESEK KOTLA TKAN INTEGRA": "РАЗРЕЗ НА КОТЕЛ TKAN INTEGRA",
    "TABELA SA DIMENZIJAMA TKAN INTEGRA SILOSA": "ТАБЛИЦА С РАЗМЕРИ НА СИЛОЗ TKAN INTEGRA",
    "POLOŽAJ KOTLOVA TKAN i TKAN INTEGRA U KOTLARNICI": "РАЗПОЛОЖЕНИЕ НА КОТЛИТЕ TKAN И TKAN INTEGRA В КОТЕЛНОТО ПОМЕЩЕНИЕ",
    "Pozicioniranje kotla u kotlarnici": "Позициониране на котела в котелното помещение",
    "KASKADNI SISTEMI": "КАСКАДНИ СИСТЕМИ",
    "DODATNA OPREMA": "ДОПЪЛНИТЕЛНО ОБОРУДВАНЕ",
    "VEĆI DNEVNI SILOS": "ПО-ГОЛЯМ ДНЕВЕН СИЛОЗ",
    "AUTOMATSKI TRANSPORT PELETA": "АВТОМАТИЧЕН ТРАНСПОРТ НА ПЕЛЕТИ",
    "Tabela 1. Raspoložive snage": "Таблица 1. Налични мощности",
    "Primena TKAN kotlova u kaskadi": "Приложение на котли TKAN в каскада",
    "Specifičnosti TKAN kaskadnih sistema": "Особености на каскадните системи TKAN",
    "Turbulatori": "Турбулатори",
    "Cevni izmenjivač": "Тръбен топлообменник",
    "Vrata za čišćenje cevnog izmenjivača i samog kotla": "Врата за почистване на тръбния топлообменник и самия котел",
    "Razvodni ormar sa automatikom": "Разпределително табло с автоматика",
    "Kanta za pepeo": "Контейнер за пепел",
    "Vrata za loženje i potpalu": "Врата за зареждане и разпалване",
    "Spirala za automatsko izbacivanje pepela iz prostora ložišta": "Шнек за автоматично извеждане на пепелта от горивната камера",
    "Ložište kotla": "Горивна камера на котела",
    "Liveni segmenti": "Лети сегменти",
    "Motor za pokretanje spirale za automatsko čišćenje ložišta": "Мотор за задвижване на шнека за автоматично почистване на горивната камера",
    "Donja osovina pužnog transportera": "Долна ос на шнековия транспортьор",
    "Ćelijski dozator (valvola)": "Клетъчен дозатор (valvola)",
    "Gornja osovina pužnog transportera": "Горна ос на шнековия транспортьор",
    "Izmenjivač termičkog osiguranja": "Топлообменник за термична защита",
    "Boca komprimovanog vazduha": "Бутилка за сгъстен въздух",
    "Impulsni elektroventil": "Импулсен електромагнитен вентил",
    "Ozida ložišta": "Облицовка на горивната камера",
    "Ožid ložišta": "Облицовка на горивната камера",
    "Centrifugalni ventilator multiciklona": "Центробежен вентилатор на мултициклона",
    "Multiciklon": "Мултициклон",
    "DIMENZIJE": "РАЗМЕРИ",
    "Dimenzije": "Размери",
    "TIP KOTLA": "ТИП КОТЕЛ",
    "Tip kotla": "Тип котел",
    "Jed.mere": "Мерна ед.",
    "Snaga": "Мощност",
    "Radni pritisak": "Работно налягане",
    "Probni pritisak": "Пробно налягане",
    "Zapremina vode u kotlu": "Обем вода в котела",
    "Masa kotla": "Маса на котела",
    "Masa silosa": "Маса на силоза",
    "Količina peleta koja staje u silosu": "Количество пелети в силоза",
    "Proizvođač predviđa:": "Производителят предвижда:",
    "pužne transportere,": "шнекови транспортьори,",
    "pogone/motore pužnih transportera,": "задвижвания/мотори на шнековите транспортьори,",
    "povezivanje velikog i dnevnog silosa,": "свързване на големия и дневния силоз,",
    "automatsko dopunjavanje dnevnog silosa,": "автоматично допълване на дневния силоз,",
    "sonde minimuma i maksimuma u dnevnom silosu.": "сонди за минимум и максимум в дневния силоз.",
    "MIN → uključi transport peleta": "MIN -> включва транспорта на пелети",
    "MAX → zaustavi transport peleta": "MAX -> спира транспорта на пелети",
    "više kotlova,": "няколко котела,",
    "duži dimovod,": "по-дълъг димоотвод,",
    "ciklon/multiciklon,": "циклон/мултициклон,",
    "veći broj kolena,": "по-голям брой колена,",
    "zajednički dimnjak.": "общ комин.",
    ".": "",
    "Radijator Inženjering je domaći proizvođač kotlova na biomasu sa dugogodišnjom tradicijom, prepoznat po pouzdanim i tehnološki naprednim rešenjima za grejanje. Razvojem sopstvenih proizvodnih procesa i primenom savremenih tehnologija, stvaramo proizvode koji odgovaraju najvišim zahtevima tržišta kada su u pitanju kvalitet, efikasnost i dug vek trajanja.": "Radijator Inženjering е местен производител на котли на биомаса с дългогодишна традиция, разпознаваем с надеждни и технологично напреднали решения за отопление. Чрез развитие на собствени производствени процеси и прилагане на съвременни технологии създаваме продукти, които отговарят на най-високите пазарни изисквания за качество, ефективност и дълъг експлоатационен живот.",
    "Naša proizvodnja obuhvata kotlove snage od 6 do 600 kW, namenjene grejanju porodičnih kuća, stambenih i poslovnih objekata, javnih ustanova i industrijskih postrojenja. Zahvaljujući širokom asortimanu i mogućnosti izrade kaskadnih sistema, u mogućnosti smo da ponudimo optimalno rešenje za svaki objekat i svaku potrebu.": "Нашето производство обхваща котли с мощност от 6 до 600 kW, предназначени за отопление на семейни къщи, жилищни и търговски обекти, обществени институции и индустриални инсталации. Благодарение на широкия асортимент и възможността за изработка на каскадни системи можем да предложим оптимално решение за всеки обект и всяка нужда.",
    "Kvalitet naših proizvoda rezultat je pažljivo odabranih materijala, precizne proizvodnje i rigorozne kontrole kvaliteta u svim fazama procesa. Svaki kotao razvijen je sa ciljem da obezbedi maksimalnu energetsku efikasnost, pouzdan rad i dugoročnu eksploataciju uz minimalne troškove održavanja.": "Качеството на нашите продукти е резултат от внимателно подбрани материали, прецизно производство и строг контрол на качеството във всички фази на процеса. Всеки котел е разработен с цел да осигури максимална енергийна ефективност, надеждна работа и дългосрочна експлоатация с минимални разходи за поддръжка.",
    "Posebnu pažnju posvećujemo inovacijama i unapređenju tehnologije proizvodnje, kako bismo korisnicima obezbedili savremena rešenja koja kombinuju visok stepen automatizacije, jednostavno upravljanje i maksimalno iskorišćenje energije.": "Обръщаме специално внимание на иновациите и усъвършенстването на производствените технологии, за да осигурим на потребителите съвременни решения, които съчетават висока степен на автоматизация, лесно управление и максимално оползотворяване на енергията.",
    "Svi naši proizvodi projektovani su i proizvedeni u skladu sa važećim evropskim standardima i propisima u oblasti kotlova, što predstavlja potvrdu njihove bezbednosti, pouzdanosti i visokog kvaliteta.": "Всички наши продукти са проектирани и произведени в съответствие с действащите европейски стандарти и разпоредби в областта на котлите, което потвърждава тяхната безопасност, надеждност и високо качество.",
    "Izborom Radijator Inženjering kotlova birate domaći proizvod, provereni kvalitet i partnera koji svojim iskustvom, stručnom podrškom i kompletnim sistemskim rešenjima pruža sigurnost tokom celog životnog veka sistema grejanja.": "Избирайки котлите Radijator Inženjering, избирате местен продукт, доказано качество и партньор, който със своя опит, професионална поддръжка и цялостни системни решения осигурява сигурност през целия жизнен цикъл на отоплителната система.",
    "Kompanija sa 35 godina iskustva u projektovanju, izradi i inovacijama na polju kotlova koji zagrevaju hiljade objekata širom Evrope!": "Компания с 35 години опит в проектирането, производството и иновациите в областта на котлите, които отопляват хиляди обекти в цяла Европа!",
    "Industrijski kotlovi na biomasu su izrađeni od kotlovskih limova debljine 6 mm i više. Izmenjivač toplote je od bešavnih, kotlovskih cevi. Stepen iskorišćenja preko 90% na pelet. Temperature dimnih gasova na izlazu su od 170 do 190 stepeni pri višim režimima, što uvek možemo da proverimo na displeju automatike. Dostupni su u opsegu snaga od 60 do 300 kW.": "Индустриалните котли на биомаса са изработени от котелна ламарина с дебелина 6 mm и повече. Топлообменникът е от безшевни котелни тръби. КПД при работа с пелети е над 90%. Температурата на димните газове на изхода е от 170 до 190 градуса при по-високи режими и винаги може да се провери на дисплея на автоматиката. Предлагат се в мощностен диапазон 60-300 kW.",
    "Industrijski kotao na biomasu predstavlja unapređenu verziju standardnog TKAN kotla. Opremljen je zidanim ložištem, naprednijom automatikom i bogatijom pratećom opremom, čime se postižu veća pouzdanost, bolja efikasnost sagorevanja i šire mogućnosti primene. Izrađen je od kotlovskih limova debljine 6 mm i više, sa izmenjivačem toplote od bešavnih kotlovskih cevi. Stepen iskorišćenja je preko 90% Dostupan je u opsegu snaga od 80 do 600 kW.": "Индустриалният котел на биомаса представлява усъвършенствана версия на стандартния котел TKAN. Оборудван е със зидана горивна камера, по-усъвършенствана автоматика и по-богато съпътстващо оборудване, с което се постигат по-висока надеждност, по-добра ефективност на горенето и по-широки възможности за приложение. Изработен е от котелна ламарина с дебелина 6 mm и повече, с топлообменник от безшевни котелни тръби. КПД е над 90%, а мощностният диапазон е 80-600 kW.",
    "Kaskadni sistemi predstavljaju kombinaciju dva ili više kotlova povezanih u jedinstven sistem sa zajedničkim silosom za skladištenje peleta. Ovakva konfiguracija omogućava veću ukupnu instalisanu snagu, pouzdaniji rad, ravnomernu raspodelu opterećenja i veću energetsku efikasnost sistema.": "Каскадните системи представляват комбинация от два или повече котела, свързани в единна система с общ силоз за съхранение на пелети. Такава конфигурация осигурява по-голяма обща инсталирана мощност, по-надеждна работа, равномерно разпределение на натоварването и по-висока енергийна ефективност на системата.",
    "Pored kotlova, u ponudi je kompletna prateća oprema za formiranje funkcionalnog sistema grejanja. Asortiman obuhvata silose za skladištenje peleta, pužne transportere, elevatore, bafer rezervoare, automatiku i ostalu opremu potrebnu za pouzdan transport, doziranje i kontrolu peleta, kao i siguran i efikasan rad celokupnog sistema.": "Освен котлите, в предложението влиза и пълно съпътстващо оборудване за изграждане на функционална отоплителна система. Асортиментът включва силози за съхранение на пелети, шнекови транспортьори, елеватори, буферни съдове, автоматика и друго оборудване, необходимо за надежден транспорт, дозиране и контрол на пелетите, както и за безопасна и ефективна работа на цялата система.",
    "Kotao TKAN je razvijen sa ciljem da RADIJATOR INŽENJERING ponudi tržištu kotao koji je po svojim mehaničkim i termičkim osobinama izrazito namenjen biomasi kao gorivu. Sa druge strane zahtevi tržišta su uvek okrenuti ka što većoj univerzalnosti goriva, tako da je TKAN moguće ložiti i sa drvetom i tada je loženje ručno.": "Котелът TKAN е разработен с цел RADIJATOR INŽENJERING да предложи на пазара котел, чиито механични и термични характеристики са специално пригодени за биомаса като гориво. От друга страна, изискванията на пазара са насочени към възможно най-голяма универсалност на горивото, поради което TKAN може да се захранва и с дърва, като тогава зареждането е ръчно.",
    "Po spoljašnjem dizajnu, dimenzijama ložišta, otvorima za loženje i čišćenje TKAN je zadržao sve dobre osobine predhodnih modela po kojima je RADIJATOR INŽENJERING prepoznatljiv na tržištu. Vodeni deo kotla, njegov način izmene toplote između dimnih gasova i vode putem cevnog izmenjivača, prilagođen je biomasi. Zbog primene ventilatora tj. prinudne promaje put dimnih gasova duži je nego kod standardnih kotlova. Iz istih razloga moguća je primena usmerivača dimnih gasova tzv. turbulatora koji dodatno povećavaju stepen iskorišćenja kotla. Turbulatori su spirale napravljene od specijalnog materijala.": "По външен дизайн, размери на горивната камера и отвори за зареждане и почистване TKAN е запазил всички добри качества на предишните модели, с които RADIJATOR INŽENJERING е разпознаваем на пазара. Водната част на котела и начинът на топлообмен между димните газове и водата чрез тръбен топлообменник са пригодени за биомаса. Поради използването на вентилатор, тоест принудителна тяга, пътят на димните газове е по-дълъг отколкото при стандартните котли. По същите причини е възможно използване на насочващи елементи за димните газове, т.нар. турбулатори, които допълнително повишават КПД на котела. Турбулаторите са спирали, изработени от специален материал.",
    "Stepen korisnosti na pelet je preko 90%. Pri normalnim režimima temperatura dimnih gasovana izlazu je oko 160 ̊C, a pri maksimalnim režimima je ispod 180 ̊C. Ove vrednosti mogu u svakom trenutku da se očitaju na displeju.": "КПД при работа с пелети е над 90%. При нормални режими температурата на димните газове на изхода е около 160 °C, а при максимални режими е под 180 °C. Тези стойности могат по всяко време да се отчетат на дисплея.",
    "Svi delovi vodenog dela kotla izrađeni su od bešavnih cevi kvaliteta ST 35.4 i kotlovskih limova debljine 5mm i više, u zavisnosti od snage kotla. Limovi su kvaliteta 1.0425 EU standard odnosno P265GH standard EUII. Ložište je po svom principu rada tzv. „izviruće“, gde gorivo iz zone transporta ide vertikalno uvis tj. izvire do zone sagorevanja. Napravljeno je od masivnih izolacijskih materijala i sivog liva. Transport goriva obezbeđen je pužnim transporterima.": "Всички части на водната част на котела са изработени от безшевни тръби с качество ST 35.4 и котелна ламарина с дебелина 5 mm и повече, в зависимост от мощността на котела. Ламарините са с качество 1.0425 по стандарт EU, съответно P265GH по стандарт EUII. Горивната камера работи на принципа на т.нар. възходящо горене, при което горивото от зоната на транспорта се движи вертикално нагоре към зоната на горене. Изработена е от масивни изолационни материали и сив чугун. Транспортът на горивото се осигурява чрез шнекови транспортьори.",
    "TKAN INTEGRA predstavlja novu generaciju industrijskih kotlova na biomasu, razvijenu kao unapređenje standardnog TKAN modela. Nastao je kao odgovor na zahteve tržišta za većim stepenom automatizacije, višom energetskom efikasnošću i većom pouzdanošću u radu, uz zadržavanje svih proverenih konstrukcionih rešenja po kojima je RADIJATOR INŽENJERING prepoznatljiv.": "TKAN INTEGRA представлява ново поколение индустриални котли на биомаса, разработено като усъвършенстване на стандартния модел TKAN. Създаден е като отговор на пазарните изисквания за по-висока степен на автоматизация, по-висока енергийна ефективност и по-голяма надеждност при работа, като запазва всички доказани конструктивни решения, с които RADIJATOR INŽENJERING е разпознаваем.",
    "Zahvaljujući zidanom ložištu, unapređenom sistemu sagorevanja, savremenoj automatici i bogatijoj standardnoj opremi, TKAN INTEGRA obezbeđuje stabilan rad, maksimalno iskorišćenje energije i dug radni vek svih ključnih komponenti. Konstrukcija kotla izrađena je od visokokvalitetnih kotlovskih limova debljine 6 mm i više, dok je cevni izmenjivač toplote izrađen od bešavnih kotlovskih cevi, čime su obezbeđeni visoka mehanička čvrstoća, pouzdanost i dugotrajnost sistema.": "Благодарение на зиданата горивна камера, усъвършенстваната система за горене, съвременната автоматика и по-богатото стандартно оборудване, TKAN INTEGRA осигурява стабилна работа, максимално оползотворяване на енергията и дълъг експлоатационен живот на всички ключови компоненти. Конструкцията на котела е изработена от висококачествени котелни ламарини с дебелина 6 mm и повече, а тръбният топлообменник е от безшевни котелни тръби, което осигурява висока механична якост, надеждност и дълготрайност на системата.",
    "Zidano ložište predstavlja jednu od ključnih prednosti modela TKAN INTEGRA. Ovakva konstrukcija omogućava potpuno i stabilno sagorevanje goriva, minimalne emisije štetnih gasova i čestica prašine, kao i maksimalno iskorišćenje energije sadržane u peletu. Istovremeno, čelični delovi kotla nisu direktno izloženi plamenu, čime se značajno produžava radni vek kotla.": "Зиданата горивна камера е едно от ключовите предимства на модела TKAN INTEGRA. Тази конструкция позволява пълно и стабилно изгаряне на горивото, минимални емисии на вредни газове и прахови частици, както и максимално оползотворяване на енергията в пелетите. Едновременно с това стоманените части на котела не са директно изложени на пламъка, което значително удължава експлоатационния живот на котела.",
    "Za smanjenje količine čestica koje odlaze u dimnjak predviđen je ciklon, odnosno kod većih konfiguracija multiciklon sa centrifugalnim ventilatorom. Proizvođač posebno preporučuje ciklon kada se koristi pneumatsko čišćenje izmenjivača, jer se tada dodatno izbacuje pepeo/čađ iz kotla. Kod većih TKAN Integra modela koristi se multiciklon sa centrifugalnim ventilatorom. Tačnije, kod modela TKAN 80 Integra i TKAN 100 Integra ugrađen je ventilator na dimnjači, dok je kod modela TKAN 150 Integra, TKAN 200 Integra, TKAN 250 Integra i TKAN 300 Integra konstruisan multiciklon sa centrifugalnim ventilatorom.": "За намаляване на количеството частици, които отиват към комина, е предвиден циклон, а при по-големи конфигурации - мултициклон с центробежен вентилатор. Производителят особено препоръчва циклон, когато се използва пневматично почистване на топлообменника, тъй като тогава допълнително се извеждат пепел и сажди от котела. При по-големите модели TKAN Integra се използва мултициклон с центробежен вентилатор. По-точно, при моделите TKAN 80 Integra и TKAN 100 Integra вентилаторът е монтиран на димохода, а при моделите TKAN 150 Integra, TKAN 200 Integra, TKAN 250 Integra и TKAN 300 Integra е конструиран мултициклон с центробежен вентилатор.",
    "Kod pojedinih TKAN konfiguracija koristi se ventilator na dimovodnoj strani, a kod većih sistema multiciklon može biti opremljen centrifugalnim ventilatorom. Ovo je bitno prilikom projektovanja kompletnog dimnjaka, naročito ako ima:": "При отделни конфигурации TKAN се използва вентилатор от страната на димоотвода, а при по-големи системи мултициклонът може да бъде оборудван с центробежен вентилатор. Това е важно при проектирането на цялата коминна система, особено ако има:",
    "Kotlovi serije TKAN INTEGRA standardno su opremljeni sistemom za automatsko čišćenje prostora oko ložišta, dok se čišćenje cevnog izmenjivača vrši automatski pomoću komprimovanog vazduha. Sistem periodičnim vazdušnim impulsima uklanja naslage čađi iz dimovodnih cevi, održava visok stepen iskorišćenja kotla i značajno smanjuje potrebu za ručnim održavanjem.": "Котлите от серия TKAN INTEGRA стандартно са оборудвани със система за автоматично почистване на пространството около горивната камера, а тръбният топлообменник се почиства автоматично със сгъстен въздух. Чрез периодични въздушни импулси системата премахва наслояванията от сажди в димогарните тръби, поддържа висок КПД на котела и значително намалява нуждата от ръчна поддръжка.",
    "Kotlarnica mora biti obezbeđena od smrzavanja. Podloga za kotao u kotlarnici mora biti od nezapaljivog materijala. Preporučene vrednosti udaljenosti sve četiri strane kotla u odnosu na zidove kotlarnice ili neka druga kruta tela (akomulacioni bojler itd.) prikazane su tablično slici ispod.": "Котелното помещение трябва да бъде защитено от замръзване. Основата за котела трябва да бъде от негорим материал. Препоръчителните разстояния от всички четири страни на котела до стените на котелното помещение или до други твърди тела (акумулиращ бойлер и др.) са показани в таблицата и изображението по-долу.",
    "Kompanija Radijator Inženjering razvila je seriju kotlova TKAN prvenstvenstveno za sagorevanje biomase (peleta, koštica voća) i drveta. Kada se ovi kotlovi povezuju u kaskadne sisteme, dobija se izuzetno moćno i fleksibilno rešenje za grejanje velikih objekata.": "Компанията Radijator Inženjering разработи серията котли TKAN преди всичко за изгаряне на биомаса (пелети, костилки от плодове) и дърва. Когато тези котли се свържат в каскадни системи, се получава изключително мощно и гъвкаво решение за отопление на големи обекти.",
    "Za kaskadne sisteme najčešće se koriste industrijski modeli veće snage, kao što su TKAN 100, 150, 200, 250 i 300 kW.": "За каскадни системи най-често се използват индустриални модели с по-голяма мощност, като TKAN 100, 150, 200, 250 и 300 kW.",
    "Pokrivanje velikih snaga: Povezivanjem npr. dva kotla TKAN 300 u kaskadu, dobija se sistem ukupne snage od 600 kW koji može da greje hotele, proizvodne hale ili stambene komplekse.": "Покриване на големи мощности: чрез свързване например на два котела TKAN 300 в каскада се получава система с обща мощност 600 kW, която може да отоплява хотели, производствени халета или жилищни комплекси.",
    "Modularnost i fleksibilnost: U prelaznim periodima (jesen/proleće) radi samo jedan TKAN kotao na optimalnom režimu, dok se drugi pali tek kada spoljna temperatura drastično padne.": "Модулност и гъвкавост: в преходните периоди (есен/пролет) работи само един котел TKAN в оптимален режим, а вторият се включва едва когато външната температура значително спадне.",
    "Upravljanje i automatika: TKAN kotlovi poseduju naprednu elektroniku koja preko spoljnih kaskadnih regulatora omogućava sinhronizovan rad. Automatika prati temperaturu u hidrauličnoj skretnici i komanduje koji će kotao startovati.": "Управление и автоматика: котлите TKAN разполагат с усъвършенствана електроника, която чрез външни каскадни регулатори позволява синхронизирана работа. Автоматиката следи температурата в хидравличния разделител и управлява кой котел да стартира.",
    "Kontinuirano snabdevanje gorivom: Industrijski TKAN kotlovi dolaze sa dnevnim silosima (npr. 800 litara) koji se preko dodatnih pužnih transportera mogu povezati sa jednim velikim, centralnim silosom za pelet koji snabdeva celu kaskadu.": "Непрекъснато снабдяване с гориво: индустриалните котли TKAN се доставят с дневни силози (например 800 литра), които чрез допълнителни шнекови транспортьори могат да се свържат с един голям централен силоз за пелети, захранващ цялата каскада.",
    "Sigurnost i kontinuitet: Ukoliko je na jednom kotlu potrebno uraditi čišćenje pepela ili redovan servis, hidraulički sistem i kaskadna automatika omogućavaju da drugi kotao nesmetano nastavi rad, tako da objekat nikada ne ostaje bez grejanja.": "Сигурност и непрекъснатост: ако на един котел е необходимо почистване на пепелта или редовно сервизиране, хидравличната система и каскадната автоматика позволяват на другия котел да продължи работа без прекъсване, така че обектът никога да не остава без отопление.",
    "Za pouzdan i efikasan rad industrijskih kotlovskih postrojenja, sistem se može opremiti dodatnim komponentama prilagođenim potrebama objekta i načinu korišćenja.": "За надеждна и ефективна работа на индустриалните котелни инсталации системата може да бъде оборудвана с допълнителни компоненти, съобразени с нуждите на обекта и начина на използване.",
    "Dodatna oprema obuhvata sisteme za automatsko doziranje i transport goriva, skladištenje peleta, automatizovanu regulaciju rada, kao i opremu za povezivanje i kaskadno upravljanje kotlovima.": "Допълнителното оборудване включва системи за автоматично дозиране и транспорт на гориво, съхранение на пелети, автоматизирано регулиране на работата, както и оборудване за свързване и каскадно управление на котлите.",
    "Rešenja se projektuju prema kapacitetu kotlarnice, potrebnoj autonomiji rada i raspoloživom prostoru, sa ciljem postizanja visokog stepena automatizacije, pouzdanosti i optimalne potrošnje goriva.": "Решенията се проектират според капацитета на котелното помещение, необходимата автономност на работа и наличното пространство, с цел постигане на висока степен на автоматизация, надеждност и оптимален разход на гориво.",
    "U ložišnom delu, za automatsko izdvajanje pepela ugrađuju se dve pužne spirale sa svojim elektro pogonima. One pepeo ubacuju u dve kutije koje povremeno treba prazniti.": "В горивната част, за автоматично извеждане на пепелта, се монтират две шнекови спирали със собствени електрозадвижвания. Те отвеждат пепелта в две кутии, които периодично трябва да се изпразват.",
    "Na vrata izmenjivačkog sklopa cevi ugrađuje se sistem elektromagnetnih ventila koji povremeno puste vazduh pod pritiskom i na taj način čiste cevi kotla od pepela i čađi. Potreban je izvor vazduha pod pritiskom određenog kapaciteta kao i automatika koja vodi ovaj proces.": "На вратата на тръбния топлообменен блок се монтира система от електромагнитни вентили, които периодично изпускат въздух под налягане и така почистват тръбите на котела от пепел и сажди. Необходим е източник на сгъстен въздух с определен капацитет, както и автоматика, която управлява този процес.",
    "Zbog smanjene emisije čestica pepela u vazduhu, preporučuje se ugradnja ciklona naročito ako je kupac ugradio i sistem pneumatskog čišćenja.": "За намаляване на емисиите на пепелни частици във въздуха се препоръчва монтаж на циклон, особено ако клиентът е инсталирал и система за пневматично почистване.",
    "Kod velikih sistema gde se dnevna potrošnja peleta kreće i od nekoliko stotina kilograma, pa do nekoliko tona, preporučuje se ugradnja velikog silosa sa kofičastim elevatorom. On je sistemom pužnih transportera vezan sa malim silosom, a ceo proces dopreme je automatizovan sa sondama minimuma i maksimuma u malom silosu.": "При големи системи, където дневният разход на пелети достига от няколкостотин килограма до няколко тона, се препоръчва голям силоз с кофичков елеватор. Той се свързва с малкия силоз чрез шнекови транспортьори, а подаването се автоматизира със сонди в малкия силоз.",
    "Standardni TKAN kotlovi opremljeni su dnevnim silosom, dok proizvođač za veće sisteme omogućava izradu većih spoljnih silosa sa posebnim dimenzijama. U zavisnosti od potreba sistema, mogu se izraditi silosi kapaciteta nekoliko desetina tona sa kofičastim elevatorom.": "Стандартните котли TKAN са оборудвани с дневен силоз, а за по-големи системи производителят позволява изработка на по-големи външни силози със специални размери. В зависимост от нуждите на системата могат да се изработят силози с капацитет няколко десетки тона с кофичков елеватор.",
    "Veliki spoljni silos povezuje se sa dnevnim silosom kotla putem pužnih transportera, čime se omogućava automatsko dopremanje dnevnog silosa putem sondi minimuma i maksimuma. Za skladištenje većih količina peleta mogu se koristiti i Jumbo vreće.": "Големият външен силоз се свързва с дневния силоз на котела чрез шнекови транспортьори, което позволява автоматично допълване на дневния силоз чрез сонди за минимум и максимум. За съхранение на по-големи количества пелети могат да се използват и Jumbo торби.",
    "Kod novijih modela TKAN 60–300, standardni dnevni silos ima zapreminu od 800 litara, uz mogućnost povezivanja sa velikim spoljnim silosom. Povezivanje može biti izvedeno bočno ili čeono, u zavisnosti od rasporeda opreme i prostornih uslova. Dakle:": "При по-новите модели TKAN 60-300 стандартният дневен силоз е с обем 800 литра, с възможност за свързване с голям външен силоз. Свързването може да бъде странично или челно, в зависимост от разположението на оборудването и пространствените условия. Схемата е:",
    "veliki centralni silos → pužni transporter → dnevni silos TKAN → puž do ložišta": "голям централен силоз -> шнеков транспортьор -> дневен силоз TKAN -> шнек до горивната камера",
    "Kod velikih sistema preporučuje se veliki silos sa kofičastim elevatorom, pužnim transporterima i automatskim dopunjavanjem dnevnog silosa putem sondi minimuma i maksimuma. Za skladištenje većih količina peleta mogu se koristiti i Jumbo vreće.": "При големи системи се препоръчва голям силоз с кофичков елеватор, шнекови транспортьори и автоматично допълване на дневния силоз чрез сонди за минимум и максимум. За съхранение на по-големи количества пелети могат да се използват и Jumbo торби.",
    "Automatika kotla može da upravlja motorom puža za dopremu iz velikog silosa.": "Автоматиката на котела може да управлява мотора на шнека за подаване от големия силоз.",
    "Za kaskadu više TKAN kotlova, ovo je posebno interesantno jer se može projektovati centralni sistem distribucije peleta.": "За каскада от няколко котела TKAN това е особено полезно, защото може да се проектира централна система за разпределение на пелетите.",
    "Kod automatskog punjenja dnevnog silosa koriste se sonde minimuma i maksimuma. Princip je:": "При автоматично пълнене на дневния силоз се използват сонди за минимум и максимум. Принципът е:",
    "Na taj način kotao sam traži gorivo iz centralnog skladišta i ne zahteva ručno dopunjavanje.": "По този начин котелът сам заявява гориво от централния склад и не изисква ръчно допълване.",
}

BG_PREFIX_TRANSLATIONS = [
    ('Radijator inženjering" d.o.o. u poslovnom smislu je', "Radijator Inženjering d.o.o. е правен наследник на занаятчийската работилница „Radijator“, основана през 1991 г., чиято основна дейност е била монтаж и поддръжка на централно отопление. Първият водогреен котел на твърдо гориво произведохме през 1985 г."),
    ("Preduzeće u današnjoj formi postoji od 2002. godine", "Предприятието в днешната си форма съществува от 2002 г. и от година на година напредва с големи стъпки, като винаги се стреми да бъде сред първите в прилагането на нови технологии, качеството на продуктите и завладяването на нови европейски пазари."),
    ("Kako smo proširivali i usavršavali proizvodnju", "С разширяването и усъвършенстването на производството достигнахме ниво, при което котлите се изработват с най-съвременни световни технологии. В областта на рязането на ламарина се открояват лазерно рязане, CNC плазмен процес и CNC щанцоване. Заваряването се извършва роботизирано, както и с автоматизирани машини. Най-добрият показател за качеството на продуктите и услугите е фактът, че производството се увеличава всяка година."),
    ('Danas "Radijator-inženjering"', "Днес Radijator Inženjering има над 350 служители, от които 40 са дипломирани машинни инженери, които ежедневно работят за усъвършенстване на качеството на продуктите."),
    ("Sigurna postojanost kvaliteta", "Постоянното качество както на продуктите, така и на дейността на фирмата е потвърдено с получаването на сертификат за система за качество ISO 9001:2008."),
]

BG_SKIP_PREFIXES = (
    "pravni naslednik zanatske radnje",
    "osnovana 1991.",
    "montaža i održavanje",
    "toplovodni kotao",
    "godine.",
    "iz godine u godinu",
    "trudeći da bude",
    "kvalitetu proizvoda",
    "došli do nivoa",
    "svetskim tehnologijama",
    "izdvajaju se",
    "CNC probijanje",
    "kao i upotrebom",
    "proizvoda i usluga",
    "proizvodnja povećava",
    "radnika od kojih",
    "rade na usavršavanju",
    "poslovanja firme",
    "sistema kvaliteta",
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
        .replace("SERIJA TKAN INTEGRA MODELI", "SERIJA TKAN INTEGRA")
        .replace("SERIJA TKAN MODELI", "SERIJA TKAN")
        .replace("Serija TKAN Integra modeli", "Serija TKAN Integra")
        .replace("Serija TKAN modeli", "Serija TKAN")
        .replace(
            "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA MODEL",
            "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM, OTPRAŠIVANJEM I CIKLONOM – TKAN INTEGRA",
        )
        .replace(
            "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM TKAN MODEL",
            "TOPLOVODNI KOTAO NA PELET SA AUTOMATSKIM NALAGANJEM TKAN",
        )
        .replace("PRESEK TKAN INTEGRA KOTLA SA OPISOM ELEMENATA", "PRESEK KOTLA TKAN INTEGRA")
        .replace("PRESEK TKAN KOTLA SA OPISOM ELEMENATA", "PRESEK KOTLA TKAN")
        .replace("POLOŽAJ TKAN OBIČNOG I TKAN INTEGRA KOTLA U KOTLARNICI", "POLOŽAJ KOTLOVA TKAN i TKAN INTEGRA U KOTLARNICI")
        .replace("POLOŽAJ TKAN OBIČNOG i TKAN INTEGRA KOTLA U KOTLARNICI", "POLOŽAJ KOTLOVA TKAN i TKAN INTEGRA U KOTLARNICI")
    )


def translate_text(text: str, language: str) -> str:
    text = normalize_catalog_text(text)
    if language == "sr":
        return text
    if language == "bg":
        if text.startswith(BG_SKIP_PREFIXES):
            return ""
        if text in BG_EXACT_TRANSLATIONS:
            return BG_EXACT_TRANSLATIONS[text]
        for prefix, translation in BG_PREFIX_TRANSLATIONS:
            if text.startswith(prefix):
                return translation
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
    color_spread = rgb.max(axis=2) - rgb.min(axis=2)
    neutral = color_spread <= 42
    light_candidate = (brightness >= 238) & neutral & (alpha > 0)
    off_white_candidate = (
        ((brightness >= 232) & (color_spread <= 58) & (alpha > 0))
        | ((brightness >= 218) & (color_spread <= 34) & (alpha > 0))
    )

    # Remove the edge-connected backdrop and isolated off-white islands
    # left inside dimension drawings, while keeping dark callouts and colored geometry.
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

    data[connected | off_white_candidate, 3] = 0
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
    for target_name, source_path in {
        **SUPPLEMENTAL_IMAGE_SOURCES,
        **EDITORIAL_PRODUCT_IMAGE_SOURCES,
    }.items():
        if not source_path.is_file():
            raise FileNotFoundError(f"Nedostaje slika za katalog: {source_path}")
        save_display_image(source_path, EDITORIAL_ASSET_DIR / target_name)


def prepare_boiler_room_position_image() -> None:
    source_path = EDITORIAL_ASSET_DIR / "boiler-room-position.png"
    target_path = EDITORIAL_ASSET_DIR / "boiler-room-position-clean.png"
    if source_path.is_file():
        save_display_image(source_path, target_path)


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
    label = {"sr": "Slika", "ro": "Imagine", "bg": "Фигура"}.get(language, "Slika")
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
            else "Производство Radijator Inženjering"
            if language == "bg"
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
    if language == "bg":
        return """
<section class="catalog-section production-spread" id="production-standards" data-section="PRO">
  <div class="production-spread__brand">
    <span>Производство</span>
  </div>
  <div class="production-spread__headline">
    <p class="production-spread__kicker">Технология и качество</p>
    <h2>Производство по съвременни европейски стандарти</h2>
  </div>
  <div class="production-spread__media">
    <figure><img src="assets/editorial/company-aerial-complex-wide.jpg" alt="Производственият комплекс Radijator Inzenjering от въздуха" /></figure>
    <figure><img src="assets/editorial/company-aerial-complex-top.jpg" alt="Заводът Radijator Inzenjering със съвременно производство" /></figure>
  </div>
  <div class="production-spread__copy">
    <p>С разширяването и усъвършенстването на производството котлите започнаха да се изработват с най-съвременни технологии: лазерно рязане, CNC плазмен процес, CNC щанцоване, роботизирано заваряване и автоматизирано заваряване.</p>
    <p>Днес Radijator Inženjering има над 350 служители, сред които 40 дипломирани машинни инженери, които ежедневно работят за подобряване на качеството на продуктите.</p>
  </div>
  <div class="production-spread__stats">
    <div><strong>350+</strong><span>служители</span></div>
    <div><strong>40</strong><span>дипл. машинни инженери</span></div>
    <div><strong>EU</strong><span>износ в 27+ държави от ЕС</span></div>
  </div>
  <p class="production-spread__footer">Технология / качество / пазар</p>
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
        '<img src="assets/editorial/boiler-room-position-clean.png" '
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


def replace_section_images(
    body_html: str,
    section_id: str,
    replacements: list[tuple[str, str]],
) -> str:
    section_start = body_html.find(f'id="{section_id}"')
    section_end = body_html.find("</section>", section_start)
    if section_start == -1 or section_end == -1:
        return body_html

    section_html = body_html[section_start:section_end]
    replacement_index = 0

    def replace_image(match: re.Match[str]) -> str:
        nonlocal replacement_index
        if replacement_index >= len(replacements):
            return match.group(0)
        filename, alt = replacements[replacement_index]
        replacement_index += 1
        return (
            f'<img src="assets/editorial/{html.escape(filename)}" '
            f'alt="{html.escape(alt)}" />'
        )

    updated_section = re.sub(r'<img src="[^"]+" alt="[^"]*" />', replace_image, section_html)
    return body_html[:section_start] + updated_section + body_html[section_end:]


def merge_split_technical_figure(body_html: str, section_id: str, image_name: str) -> str:
    section_start = body_html.find(f'id="{section_id}"')
    section_end = body_html.find("</section>", section_start)
    if section_start == -1 or section_end == -1:
        return body_html

    section_html = body_html[section_start:section_end]
    image_pos = section_html.find(image_name)
    if image_pos == -1:
        return body_html

    container_start = section_html.rfind('<div class="technical-visual"', 0, image_pos)
    figure_start = section_html.rfind("<figure", container_start, image_pos)
    figure_end = section_html.find("</figure>", image_pos)
    container_end = section_html.find("</div></div>", figure_end)
    previous_container_end = section_html.rfind("</div></div>", 0, container_start)
    if min(container_start, figure_start, figure_end, container_end, previous_container_end) == -1:
        return body_html

    # If the figure has already been merged into the previous row, leave the section untouched.
    if section_html.find(image_name, 0, container_start) != -1:
        return body_html

    figure_end += len("</figure>")
    container_end += len("</div></div>")
    figure_html = section_html[figure_start:figure_end]
    updated_section = (
        section_html[:previous_container_end]
        + "\n    "
        + figure_html
        + "\n    "
        + section_html[previous_container_end:container_start]
        + section_html[container_end:]
    )
    return body_html[:section_start] + updated_section + body_html[section_end:]


def tune_catalog_layout(body_html: str) -> str:
    """Apply editorial moves that keep generated content aligned with the catalog story."""
    body_html = replace_section_images(
        body_html,
        "section-02",
        [
            ("tkan-300-silos.png", "TKAN 300 sa silosom"),
            ("tkan-integra-render.png", "TKAN Integra kotao"),
            ("kaskadni-sistem-render.png", "Kaskadni sistem kotlova"),
        ],
    )
    body_html = replace_section_images(
        body_html,
        "section-03",
        [("tkan-300-silos.png", "TKAN 300 sa silosom")],
    )
    body_html = replace_section_images(
        body_html,
        "section-07",
        [("tkan-integra-render.png", "TKAN Integra kotao")],
    )
    body_html = replace_section_images(
        body_html,
        "section-11",
        [("kaskadni-sistem-render.png", "Kaskadni sistem kotlova")],
    )

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

    body_html = merge_split_technical_figure(body_html, "section-12", "catalog-image-14.png")

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
    prepare_boiler_room_position_image()
    images = extract_images()
    images_by_key = {item["key"]: item for item in images}
    body_html, toc, _, _, _ = build_content(images_by_key, language)
    body_html = tune_catalog_layout(body_html)
    if language == "ro":
        body_html = tune_romanian_catalog_html(body_html)
    if language == "bg":
        body_html = (
            body_html.replace('alt="TKAN 300 sa silosom"', 'alt="TKAN 300 със силоз"')
            .replace('alt="TKAN Integra kotao"', 'alt="Котел TKAN Integra"')
            .replace('alt="Kaskadni sistem kotlova"', 'alt="Каскадна система от котли"')
            .replace('alt="Valvola TKAN 150 - presek"', 'alt="Valvola TKAN 150 - разрез"')
            .replace('alt="Valvola TKAN 300 Integra - presek"', 'alt="Valvola TKAN 300 Integra - разрез"')
            .replace('alt="Multiciklon TKAN 300"', 'alt="Мултициклон TKAN 300"')
            .replace('alt="Pozicioniranje TKAN kotla u kotlarnici"', 'alt="Позициониране на котел TKAN в котелното помещение"')
            .replace('alt="Instalirani industrijski kotao Radijator u kotlarnici"', 'alt="Инсталиран индустриален котел Radijator в котелно помещение"')
            .replace('alt="Kaskadno postrojenje sa industrijskim kotlovima Radijator"', 'alt="Каскадна инсталация с индустриални котли Radijator"')
        )
    production_spread = render_production_spread(language)
    body_html = body_html.replace("</section>", f"</section>\n{production_spread}", 1)
    toc.insert(1, f'<a href="#production-standards">{html.escape(config["production_toc"])}</a>')
    language_switch = render_language_switch(language)
    nav_pdf_link = render_pdf_link("web-nav-pdf", config["pdf_label"], config["pdf_filename"])
    hero_pdf_link = render_pdf_link("action-secondary", config["hero_secondary"], config["pdf_filename"])
    print_title_override = ""
    if language in {"ro", "bg"}:
        print_header = (
            "CAZANE INDUSTRIALE PE BIOMASĂ"
            if language == "ro"
            else "ИНДУСТРИАЛНИ КОТЛИ НА БИОМАСА"
        )
        print_title_override = f"""
    <style>
      @media print {{
        @page {{
          @top-center {{ content: "{print_header}"; }}
        }}
        @page catalog {{
          @top-center {{ content: "{print_header}"; }}
        }}
      }}
    </style>"""

    page = f"""<!doctype html>
<html lang="{config["html_lang"]}">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{html.escape(config["page_title"])}</title>
    <link rel="icon" href="assets/favicon.svg" />
    <link rel="stylesheet" href="styles.css?v=20260825-company-fit" />
    <link rel="stylesheet" href="catalog-premium.css?v=20260825-company-fit" />
    <link rel="stylesheet" href="catalog-web.css?v=20260825-company-fit" />
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
    if language == "bg":
        page = (
            page.replace('alt="Radijator Inženjering factory complex in Kraljevo"', 'alt="Производствен комплекс Radijator Inženjering в Кралево"')
            .replace('alt="Automated sheet metal processing in production"', 'alt="Автоматизирана обработка на ламарина в производството"')
            .replace('alt="Laser cutting of boiler sheet metal"', 'alt="Лазерно рязане на котелна ламарина"')
            .replace('alt="Operator supervising a modern production process"', 'alt="Оператор, който наблюдава съвременен производствен процес"')
            .replace('alt="Welding industrial boiler components"', 'alt="Заваряване на компоненти за индустриални котли"')
            .replace('alt="Instalirani industrijski kotao Radijator u kotlarnici"', 'alt="Инсталиран индустриален котел Radijator в котелно помещение"')
            .replace('alt="Kaskadno postrojenje sa industrijskim kotlovima Radijator"', 'alt="Каскадна инсталация с индустриални котли Radijator"')
        )
    config["html_path"].write_text(page, encoding="utf-8")
    config["alias_path"].write_text(page, encoding="utf-8")


if __name__ == "__main__":
    for selected_language in LANGUAGE_CONFIG:
        render_page(selected_language)
