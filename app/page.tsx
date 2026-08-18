const highlights = [
  "35+ godina iskustva u razvoju i proizvodnji kotlova",
  "350+ zaposlenih i tim od 40 diplomiranih masinskih inzenjera",
  "Opseg proizvoda od 15 do 500 kW za objekte razlicitih namena",
];

const programCards = [
  {
    title: "TKAN serija",
    range: "60 - 300 kW",
    description:
      "Toplovodni kotlovi na pelet sa automatskim nalaganjem, projektovani za visok stepen iskoriscenja, pouzdan rad i mogucnost lozenja drveta uz rucno punjenje.",
  },
  {
    title: "TKAN Integra",
    range: "80 - 500 kW",
    description:
      "Nova generacija industrijskih kotlova sa zidanim lozistem, automatskim ciscenjem, naprednom automatikom i multiciklonom za cistiji i stabilniji rad.",
  },
  {
    title: "Kaskadna resenja",
    range: "Modularna snaga",
    description:
      "Povezivanje dva ili vise kotlova u jedinstven sistem sa zajednickim silosom, ravnomernim opterecenjem i visokim stepenom energetske efikasnosti.",
  },
];

const tkanFeatures = [
  "Stepen korisnosti na pelet preko 90%",
  "Kotlovski limovi od 5 mm i vise, besavne cevi ST 35.4",
  "Temperatura dimnih gasova oko 160 C u normalnom rezimu",
  "Automatsko doziranje peleta puznim transporterima",
];

const integraFeatures = [
  "Zidano loziste za potpuno i stabilno sagorevanje",
  "Automatsko ciscenje izmenjivaca komprimovanim vazduhom",
  "Multiciklon sa centrifugalnim ventilatorom kod jacih modela",
  "Manje emisije prasine i manja potreba za rucnim odrzavanjem",
];

const cascadeBenefits = [
  "Pokrivanje velikih snaga za hotele, hale i stambene komplekse",
  "Optimizovan rad u prelaznim periodima uz paljenje po potrebi",
  "Kontinuitet rada sistema i tokom servisa jednog od kotlova",
  "Centralno snabdevanje peletom i sinhronizovana automatika",
];

const equipmentItems = [
  "Veliki spoljni silosi i dnevni silosi kapaciteta 800 litara i vise",
  "Puzni transporteri, pogoni i automatsko dopunjavanje preko sondi",
  "Automatsko izbacivanje pepela iz lozista u kontejnere",
  "Pneumatsko ciscenje izmenjivaca, ciklon i multiciklon resenja",
];

export default function Home() {
  return (
    <main className="site-shell">
      <section className="hero-section presentation-section" id="pocetak">
        <div className="hero-backdrop" aria-hidden="true" />
        <div className="hero-grid">
          <div className="hero-copy">
            <p className="eyebrow">Industrijski katalog</p>
            <h1>Biomasa, automatika i snaga za sisteme koji traze sigurnost.</h1>
            <p className="lead">
              Radijator Inzenjering razvija industrijske kotlove i kompletna
              sistemska resenja za objekte koji zahtevaju efikasno, stabilno i
              dugotrajno grejanje.
            </p>
            <div className="hero-actions">
              <a href="#program" className="button button-primary">
                Pogledaj program
              </a>
              <a href="#pdf-izvoz" className="button button-secondary">
                Priprema za PDF
              </a>
            </div>
          </div>

          <div className="hero-panel">
            <div className="stat-block">
              <span className="stat-value">15 - 500 kW</span>
              <span className="stat-label">Opseg resenja za razlicite objekte</span>
            </div>
            <div className="stat-block">
              <span className="stat-value">ISO 9001:2008</span>
              <span className="stat-label">Potvrda sistema kvaliteta i procesa</span>
            </div>
            <div className="stat-block">
              <span className="stat-value">Evropsko trziste</span>
              <span className="stat-label">
                Proizvodnja uskladjena sa vazecim evropskim standardima
              </span>
            </div>
          </div>
        </div>
      </section>

      <section className="content-band presentation-section" id="o-nama">
        <div className="section-heading">
          <p className="eyebrow">O kompaniji</p>
          <h2>Domaci proizvodjac sa ozbiljnom proizvodnjom i inzenjerskom bazom.</h2>
        </div>
        <div className="story-layout">
          <div className="story-copy">
            <p>
              Kompanija je nastala iz tradicije koja pocinje 1991. godine, dok
              prvi toplovodni kotao na cvrsto gorivo datira jos iz 1985. Danas
              Radijator Inzenjering objedinjuje projektovanje, proizvodnju i
              razvoj kotlova koji greju hiljade objekata sirom Evrope.
            </p>
            <p>
              Proizvodnja se oslanja na savremene procese poput laserskog
              secenja, CNC plazme, CNC probijanja i robotskog zavarivanja, sa
              fokusom na preciznost, dug vek trajanja i dosledan kvalitet.
            </p>
          </div>
          <div className="feature-stack">
            {highlights.map((item, index) => (
              <article className="feature-card" key={item}>
                <span className="feature-index">0{index + 1}</span>
                <p>{item}</p>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="program-section presentation-section" id="program">
        <div className="section-heading">
          <p className="eyebrow">Proizvodni program</p>
          <h2>Tri jasno definisana pravca za industrijske sisteme grejanja.</h2>
        </div>
        <div className="program-grid">
          {programCards.map((card) => (
            <article className="program-card" key={card.title}>
              <p className="program-range">{card.range}</p>
              <h3>{card.title}</h3>
              <p>{card.description}</p>
            </article>
          ))}
        </div>
      </section>

      <section className="detail-section presentation-section" id="tkan">
        <div className="detail-copy">
          <p className="eyebrow">Serija TKAN</p>
          <h2>Standard industrijske pouzdanosti za pelet, uz fleksibilnost za drvo.</h2>
          <p>
            TKAN je razvijen kao kotao izrazito prilagodjen biomasi, sa cevnim
            izmenjivacem toplote, produzenim putem dimnih gasova i turbulatorima
            koji dodatno podizu iskoriscenje sistema.
          </p>
        </div>
        <div className="bullet-panel">
          {tkanFeatures.map((item) => (
            <div className="bullet-item" key={item}>
              <span />
              <p>{item}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="detail-section presentation-section inverse" id="integra">
        <div className="detail-copy">
          <p className="eyebrow">Serija TKAN Integra</p>
          <h2>Naprednija automatizacija, cistije sagorevanje i duzi radni vek.</h2>
          <p>
            Integra model donosi zidano loziste, bogatiju standardnu opremu i
            sistem automatskog ciscenja koji znacajno smanjuje servisna
            opterecenja i odrzava visok stepen energetske efikasnosti.
          </p>
        </div>
        <div className="bullet-panel">
          {integraFeatures.map((item) => (
            <div className="bullet-item" key={item}>
              <span />
              <p>{item}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="content-band presentation-section" id="kaskada">
        <div className="section-heading">
          <p className="eyebrow">Kaskadni sistemi</p>
          <h2>Modularna arhitektura za velike objekte i neprekidan rad.</h2>
        </div>
        <div className="two-column-panel">
          <div className="panel-card emphasis">
            <p>
              Povezivanjem vise TKAN kotlova dobija se fleksibilan sistem koji
              precizno odgovara stvarnom opterecenju objekta, uz hidraulicnu
              skretnicu, bafer i centralizovano upravljanje.
            </p>
          </div>
          <div className="panel-card">
            {cascadeBenefits.map((item) => (
              <div className="bullet-item compact" key={item}>
                <span />
                <p>{item}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="content-band presentation-section" id="oprema">
        <div className="section-heading">
          <p className="eyebrow">Dodatna oprema</p>
          <h2>Kompletan ekosistem za skladistenje, transport i kontrolu peleta.</h2>
        </div>
        <div className="equipment-grid">
          {equipmentItems.map((item) => (
            <article className="equipment-card" key={item}>
              <p>{item}</p>
            </article>
          ))}
        </div>
      </section>

      <section className="final-section presentation-section" id="pdf-izvoz">
        <div className="final-panel">
          <p className="eyebrow">Web + PDF</p>
          <h2>Jedna prezentacija, dva nacina koriscenja.</h2>
          <p>
            Stranica je pripremljena kao elegantna web prezentacija, ali i sa
            print stilovima za cist izvoz u PDF kroz browser opciju{" "}
            <strong>Print / Save as PDF</strong>.
          </p>
          <p>
            Sledeci korak je dopuna sa tehnickim tabelama, konkretnim modelima i
            kontakt sekcijom ukoliko zelis da prezentacija ide direktno prema
            kupcima, investitorima ili distributerima.
          </p>
        </div>
      </section>
    </main>
  );
}
