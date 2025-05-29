document.addEventListener('DOMContentLoaded', function() {
  const exportBtn = document.getElementById('exportBtn');
  const status = document.getElementById('status');
  const progress = document.getElementById('progress');
  const progressBar = document.getElementById('progressBar');

  exportBtn.addEventListener('click', async () => {
    try {
      exportBtn.disabled = true;
      status.textContent = 'Authenticating...';
      progress.style.display = 'block';
      progressBar.style.width = '10%';

      // Send message to background script to start export
      chrome.runtime.sendMessage({ action: 'startExport' }, (response) => {
        if (chrome.runtime.lastError) {
          status.textContent = `Error: ${chrome.runtime.lastError.message}`;
          exportBtn.disabled = false;
          progress.style.display = 'none';
          return;
        }
        if (response && response.error) {
          status.textContent = `Error: ${typeof response.error === 'object' ? response.error.message || JSON.stringify(response.error) : response.error}`;
          exportBtn.disabled = false;
          progress.style.display = 'none';
          return;
        }
      });
    } catch (error) {
      status.textContent = `Error: ${error.message || error.toString()}`;
      exportBtn.disabled = false;
      progress.style.display = 'none';
    }
  });

  // Listen for progress updates from background script
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.type === 'progress') {
      progressBar.style.width = `${message.progress}%`;
      status.textContent = message.status;
    } else if (message.type === 'complete') {
      status.textContent = 'Export complete!';
      exportBtn.disabled = false;
      progress.style.display = 'none';
    } else if (message.type === 'error') {
      status.textContent = `Error: ${typeof message.error === 'object' ? message.error.message || JSON.stringify(message.error) : message.error}`;
      exportBtn.disabled = false;
      progress.style.display = 'none';
    }
  });
}); 