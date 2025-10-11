(function () {
  const form = document.querySelector('#feedback-form');
  const list = document.querySelector('#feedback-list');

  if (!form || !list) return;

  const STORAGE_KEY = 'ragBookFeedback';
  let feedbackItems = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');

  const renderFeedback = () => {
    list.innerHTML = '';

    if (!feedbackItems.length) {
      list.innerHTML = '<p class="highlight">Be the first to leave your reflections on the book!</p>';
      return;
    }

    [...feedbackItems]
      .sort((a, b) => b.timestamp - a.timestamp)
      .forEach(({ name, role, message, timestamp }) => {
        const entry = document.createElement('article');
        entry.className = 'feedback-item';
        entry.innerHTML = `
          <h4>${name}${role ? ` · <span>${role}</span>` : ''}</h4>
          <span>${new Date(timestamp).toLocaleString()}</span>
          <p>${message}</p>
        `;
        list.appendChild(entry);
      });
  };

  renderFeedback();

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const entry = {
      name: (data.get('name') || 'Anonymous').trim(),
      role: (data.get('role') || '').trim(),
      message: (data.get('message') || '').trim(),
      timestamp: Date.now(),
    };

    if (!entry.message) {
      alert('Please share a few words of feedback.');
      return;
    }

    feedbackItems = [...feedbackItems, entry];
    localStorage.setItem(STORAGE_KEY, JSON.stringify(feedbackItems));
    renderFeedback();
    form.reset();
  });
})();
