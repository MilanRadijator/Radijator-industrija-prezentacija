const highlights = [
  "35+ years of experience in boiler development and manufacturing",
  "350+ employees and a team of 40 graduate mechanical engineers",
  "Product range from 15 to 500 kW for buildings with different uses",
];

const programCards = [
  {
    title: "TKAN series",
    range: "60 - 300 kW",
    description:
      "Hot water pellet boilers with automatic fuel feed, designed for high efficiency and reliable operation, with the option of burning wood through manual loading.",
  },
  {
    title: "TKAN Integra",
    range: "80 - 500 kW",
    description:
      "A new generation of industrial boilers with a refractory-lined combustion chamber, automatic cleaning, advanced controls and a multicyclone for cleaner and more stable operation.",
  },
  {
    title: "Cascade solutions",
    range: "Modular output",
    description:
      "Two or more boilers connected in a single system with a shared silo, balanced loading and high energy efficiency.",
  },
];

const tkanFeatures = [
  "Efficiency above 90% when burning pellets",
  "Boiler steel plates at least 5 mm thick and seamless ST 35.4 tubes",
  "Flue gas temperature approximately 160 °C under normal operating conditions",
  "Automatic pellet dosing using screw conveyors",
];

const integraFeatures = [
  "Refractory-lined combustion chamber for complete and stable combustion",
  "Automatic heat exchanger cleaning using compressed air",
  "Multicyclone with a centrifugal fan on higher-output models",
  "Lower dust emissions and less need for manual maintenance",
];

const cascadeBenefits = [
  "High output for hotels, industrial halls and residential complexes",
  "Optimised operation in autumn and spring, with boilers starting as needed",
  "Continuous system operation while one boiler undergoes servicing",
  "Central pellet supply and synchronised controls",
];

const equipmentItems = [
  "Large external silos and day silos with capacities of 800 litres or more",
  "Screw conveyors, drives and automatic refilling controlled by level sensors",
  "Automatic ash removal from the combustion chamber into containers",
  "Pneumatic heat exchanger cleaning, cyclone and multicyclone solutions",
];

export default function Home() {
  return (
    <main className="site-shell">
      <section className="hero-section presentation-section" id="pocetak">
        <div className="hero-backdrop" aria-hidden="true" />
        <div className="hero-grid">
          <div className="hero-copy">
            <p className="eyebrow">Industrial catalogue</p>
            <h1>Biomass, automation and output for systems that demand reliability.</h1>
            <p className="lead">
              Radijator Inzenjering develops industrial boilers and complete
              system solutions for buildings that require efficient, stable
              heating over a long service life.
            </p>
            <div className="hero-actions">
              <a href="#program" className="button button-primary">
                View the range
              </a>
              <a href="#pdf-izvoz" className="button button-secondary">
                PDF preparation
              </a>
            </div>
          </div>

          <div className="hero-panel">
            <div className="stat-block">
              <span className="stat-value">15 - 500 kW</span>
              <span className="stat-label">Solutions for different building types</span>
            </div>
            <div className="stat-block">
              <span className="stat-value">ISO 9001:2008</span>
              <span className="stat-label">Quality management and process certification</span>
            </div>
            <div className="stat-block">
              <span className="stat-value">European market</span>
              <span className="stat-label">
                Production complies with applicable European standards
              </span>
            </div>
          </div>
        </div>
      </section>

      <section className="content-band presentation-section" id="o-nama">
        <div className="section-heading">
          <p className="eyebrow">About the company</p>
          <h2>A Serbian manufacturer with substantial production and engineering capabilities.</h2>
        </div>
        <div className="story-layout">
          <div className="story-copy">
            <p>
              The company grew from a tradition dating back to 1991, while its
              first hot water boiler for solid fuel dates back to 1985. Today,
              Radijator Inzenjering brings together the design, manufacturing
              and development of boilers that heat thousands of buildings across Europe.
            </p>
            <p>
              Production uses modern processes such as laser cutting, CNC plasma
              cutting, CNC punching and robotic welding, with a focus on
              precision, long service life and consistent quality.
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
          <p className="eyebrow">Product range</p>
          <h2>Three clearly defined approaches to industrial heating systems.</h2>
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
          <p className="eyebrow">TKAN series</p>
          <h2>Industrial reliability with pellets and the flexibility to burn wood.</h2>
          <p>
            TKAN is a boiler specifically adapted to biomass, with a tubular
            heat exchanger, an extended flue gas path and turbulators
            that further improve system efficiency.
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
          <p className="eyebrow">TKAN Integra series</p>
          <h2>More advanced automation, cleaner combustion and longer service life.</h2>
          <p>
            The Integra model includes a refractory-lined combustion chamber,
            more extensive standard equipment and an automatic cleaning system
            that significantly reduces maintenance work and maintains high energy efficiency.
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
          <p className="eyebrow">Cascade systems</p>
          <h2>Modular architecture for large buildings and continuous operation.</h2>
        </div>
        <div className="two-column-panel">
          <div className="panel-card emphasis">
            <p>
              Connecting several TKAN boilers creates a flexible system that
              precisely matches the building&apos;s actual heating load, with a hydraulic
              separator, buffer tank and centralised controls.
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
          <p className="eyebrow">Additional equipment</p>
          <h2>A complete system for pellet storage, transport and control.</h2>
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
          <h2>One presentation, two ways to use it.</h2>
          <p>
            The page serves as an elegant web presentation and includes print
            styles for a clean PDF export using the browser&apos;s{" "}
            <strong>Print / Save as PDF</strong>.
          </p>
          <p>
            The next step is to add technical tables, specific models and a
            contact section if you want to send the presentation directly to
            customers, investors or distributors.
          </p>
        </div>
      </section>
    </main>
  );
}
