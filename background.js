chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) return

  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['retool-importer.js'],
    })
  } catch (err) {
    console.error('Failed to inject Retool importer:', err)
  }
})
