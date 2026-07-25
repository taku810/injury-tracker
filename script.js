const revealEls = document.querySelectorAll('.reveal');

const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      observer.unobserve(entry.target);
    }
  });
}, { threshold: 0.15 });

revealEls.forEach((el) => observer.observe(el));

const navToggle = document.getElementById('navToggle');
const navPanel = document.getElementById('navPanel');

if (navToggle && navPanel) {
  const closeNav = () => {
    navToggle.classList.remove('is-open');
    navPanel.classList.remove('is-open');
    navToggle.setAttribute('aria-expanded', 'false');
  };

  navToggle.addEventListener('click', () => {
    const isOpen = navPanel.classList.toggle('is-open');
    navToggle.classList.toggle('is-open', isOpen);
    navToggle.setAttribute('aria-expanded', String(isOpen));
  });

  navPanel.querySelectorAll('a').forEach((link) => {
    link.addEventListener('click', closeNav);
  });

  document.addEventListener('click', (e) => {
    if (!navPanel.contains(e.target) && !navToggle.contains(e.target)) {
      closeNav();
    }
  });
}
