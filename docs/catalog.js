(() => {
  const root = document.documentElement;
  const body = document.body;
  const nav = document.querySelector(".catalog-web-nav");
  const progress = document.querySelector(".catalog-progress span");
  const toc = document.querySelector(".catalog-toc");
  const sections = [...document.querySelectorAll(".catalog-section")];
  const tocLinks = [...document.querySelectorAll('.catalog-toc a[href^="#section-"]')];
  const motionAllowed = !window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  root.classList.add("catalog-has-js");

  if (window.matchMedia("(max-width: 960px)").matches && toc) {
    toc.open = false;
  }

  const updateScroll = () => {
    const distance = document.documentElement.scrollHeight - window.innerHeight;
    const ratio = distance > 0 ? Math.min(window.scrollY / distance, 1) : 0;
    progress?.style.setProperty("transform", `scaleX(${ratio})`);
    nav?.classList.toggle("is-scrolled", window.scrollY > 48);
  };

  let scrollFrame = 0;
  window.addEventListener("scroll", () => {
    if (scrollFrame) return;
    scrollFrame = window.requestAnimationFrame(() => {
      updateScroll();
      scrollFrame = 0;
    });
  }, { passive: true });
  updateScroll();

  if ("IntersectionObserver" in window) {
    const revealObserver = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        entry.target.classList.add("is-visible");
        revealObserver.unobserve(entry.target);
      });
    }, { rootMargin: "0px 0px -8%", threshold: 0.08 });

    sections.forEach((section, index) => {
      section.classList.add("catalog-reveal");
      section.style.setProperty("--reveal-order", String(index % 3));
      revealObserver.observe(section);
    });

    const sectionObserver = new IntersectionObserver((entries) => {
      const visible = entries
        .filter((entry) => entry.isIntersecting)
        .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
      if (!visible) return;

      tocLinks.forEach((link) => {
        const active = link.getAttribute("href") === `#${visible.target.id}`;
        link.classList.toggle("is-active", active);
        if (active) link.setAttribute("aria-current", "location");
        else link.removeAttribute("aria-current");
      });
    }, { rootMargin: "-22% 0px -54%", threshold: [0.05, 0.2, 0.5] });

    sections.forEach((section) => sectionObserver.observe(section));
  } else {
    sections.forEach((section) => section.classList.add("is-visible"));
  }

  document.querySelectorAll(".table-scroll").forEach((wrapper) => {
    wrapper.tabIndex = 0;
    wrapper.setAttribute("role", "region");
    wrapper.setAttribute("aria-label", "Tabela tehničkih podataka, horizontalno skrolovanje");
  });

  if (motionAllowed && window.matchMedia("(pointer: fine)").matches) {
    document.querySelectorAll(".catalog-figure").forEach((figure) => {
      figure.addEventListener("pointermove", (event) => {
        const bounds = figure.getBoundingClientRect();
        const x = (event.clientX - bounds.left) / bounds.width - 0.5;
        const y = (event.clientY - bounds.top) / bounds.height - 0.5;
        figure.style.setProperty("--tilt-x", `${(-y * 3).toFixed(2)}deg`);
        figure.style.setProperty("--tilt-y", `${(x * 4).toFixed(2)}deg`);
      });
      figure.addEventListener("pointerleave", () => {
        figure.style.removeProperty("--tilt-x");
        figure.style.removeProperty("--tilt-y");
      });
    });

    body.addEventListener("pointermove", (event) => {
      root.style.setProperty("--pointer-x", `${(event.clientX / window.innerWidth - 0.5) * 2}`);
      root.style.setProperty("--pointer-y", `${(event.clientY / window.innerHeight - 0.5) * 2}`);
    }, { passive: true });
  }

  const lightbox = document.querySelector(".catalog-lightbox");
  const lightboxImage = lightbox?.querySelector("img");
  const lightboxCaption = lightbox?.querySelector("p");
  const closeButton = lightbox?.querySelector(".lightbox-close");

  const closeLightbox = () => {
    if (lightbox?.open) lightbox.close();
  };

  document.querySelectorAll(".catalog-figure img").forEach((image, index) => {
    image.tabIndex = 0;
    image.setAttribute("role", "button");
    image.setAttribute("aria-label", `${image.alt}. Otvori uvećani prikaz.`);

    const open = () => {
      if (!lightbox || !lightboxImage || !lightboxCaption) return;
      lightboxImage.src = image.currentSrc || image.src;
      lightboxImage.alt = image.alt;
      lightboxCaption.textContent = `Tehnički prikaz ${String(index + 1).padStart(2, "0")}`;
      lightbox.showModal();
    };

    image.addEventListener("click", open);
    image.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        open();
      }
    });
  });

  closeButton?.addEventListener("click", closeLightbox);
  lightbox?.addEventListener("click", (event) => {
    if (event.target === lightbox) closeLightbox();
  });
  lightbox?.addEventListener("close", () => {
    lightboxImage?.removeAttribute("src");
  });
})();
