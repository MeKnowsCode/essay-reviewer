// ============================================================
// Essay Reviewer — taskpane.js
// Word Add-in with Tracked Changes + Comments
// ============================================================

'use strict';

let intensity = 'balanced';
let apiKey = '';
let styleExamples = '';

// ── Init ──────────────────────────────────────────────────

Office.onReady(function(info) {
  if (info.host === Office.HostType.Word) {
    loadStoredSettings();
    loadDocumentInfo();
  }
});

function loadStoredSettings() {
  try {
    // Use Office roamingSettings for API key — encrypted, tied to Office account
    // Falls back to localStorage for style examples (non-sensitive)
    const settings = Office.context.roamingSettings;
    const storedKey = settings.get('essay_reviewer_api_key') || '';
    const storedStyle = localStorage.getItem('essay_reviewer_style') || '';

    apiKey = storedKey;
    styleExamples = storedStyle;

    if (storedKey) {
      document.getElementById('keyStatus').textContent = '✅';
      document.getElementById('apiKeyInput').placeholder = 'Key saved — paste new key to update';
    }
    if (storedStyle) {
      document.getElementById('styleExamples').value = storedStyle;
    }
  } catch(e) {
    console.log('Storage error:', e);
  }
}

function loadDocumentInfo() {
  Word.run(function(context) {
    const body = context.document.body;
    body.load('text');
    return context.sync().then(function() {
      const text = body.text || '';
      const words = text.split(/\s+/).filter(Boolean).length;
      const dot = document.getElementById('docDot');
      const docText = document.getElementById('docText');
      if (!text.trim()) {
        dot.className = 'doc-dot empty';
        // Safe: no user data used here
        docText.textContent = '';
        const strong = document.createElement('strong');
        strong.textContent = 'Empty document';
        docText.appendChild(strong);
        docText.appendChild(document.createTextNode(' — paste an essay to begin'));
      } else {
        // Safe: words is a number, not user-controlled text
        docText.textContent = '';
        const strong = document.createElement('strong');
        strong.textContent = words.toLocaleString() + ' words';
        docText.appendChild(strong);
        docText.appendChild(document.createTextNode(' · ready to review'));
      }
    });
  }).catch(function(e) {
    document.getElementById('docText').textContent = 'Could not read document';
  });
}

// ── Settings ──────────────────────────────────────────────

function saveApiKey() {
  const val = document.getElementById('apiKeyInput').value.trim();
  if (!val) { showStatus('error', '❌ Please enter your API key.'); return; }
  
  // Validate key format before saving
  if (!val.startsWith('sk-ant-') || val.length < 20) {
    showStatus('error', '❌ Invalid API key format. Keys start with sk-ant-');
    return;
  }

  apiKey = val;

  // Store in Office roamingSettings — more secure than localStorage
  // Encrypted and tied to the user's Office/Microsoft account
  try {
    const settings = Office.context.roamingSettings;
    settings.set('essay_reviewer_api_key', val);
    settings.saveAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById('keyStatus').textContent = '✅';
        document.getElementById('apiKeyInput').value = '';
        document.getElementById('apiKeyInput').placeholder = 'Key saved — paste new key to update';
        showStatus('success', '✅ API key saved securely.');
      } else {
        showStatus('error', '❌ Could not save key. Please try again.');
      }
    });
  } catch(e) {
    // Fallback to localStorage if roamingSettings unavailable
    localStorage.setItem('essay_reviewer_api_key', val);
    document.getElementById('keyStatus').textContent = '✅';
    document.getElementById('apiKeyInput').value = '';
    document.getElementById('apiKeyInput').placeholder = 'Key saved — paste new key to update';
    showStatus('success', '✅ API key saved.');
  }
}

function openStyleModal() {
  document.getElementById('styleOverlay').classList.add('open');
}

function closeStyleModal() {
  document.getElementById('styleOverlay').classList.remove('open');
}

function saveStyle() {
  const val = document.getElementById('styleExamples').value.trim();
  styleExamples = val;
  localStorage.setItem('essay_reviewer_style', val);
  const msg = document.getElementById('styleSaved');
  msg.style.display = 'block';
  setTimeout(function() {
    msg.style.display = 'none';
    closeStyleModal();
  }, 1500);
}

// ── UI Helpers ─────────────────────────────────────────────

function setIntensity(el) {
  document.querySelectorAll('.itab').forEach(function(t) { t.classList.remove('active'); });
  el.classList.add('active');
  intensity = el.dataset.val;
}

function toggleChip(label) {
  setTimeout(function() {
    const cb = label.querySelector('input');
    label.classList.toggle('on', cb.checked);
  }, 0);
}

function getFocusAreas() {
  return Array.from(document.querySelectorAll('.fchip input:checked')).map(function(cb) { return cb.value; });
}

function showStatus(type, msg) {
  const el = document.getElementById('status');
  el.className = 'status ' + type;
  if (type === 'loading') {
    el.innerHTML = '<div class="spinner"></div><span>' + msg + '</span>';
  } else {
    el.textContent = msg;
  }
  el.style.display = type === 'loading' ? 'flex' : 'block';
}

// ── Main Review ────────────────────────────────────────────

async function startReview() {
  if (!apiKey) {
    showStatus('error', '❌ Please save your Anthropic API key first.');
    return;
  }

  const btn = document.getElementById('reviewBtn');
  const prog = document.getElementById('progressWrap');
  const bar = document.getElementById('progressBar');

  btn.disabled = true;
  btn.textContent = 'Reviewing…';
  showStatus('loading', 'Reading essay and generating feedback…');

  prog.style.display = 'block';
  bar.style.animation = 'none';
  bar.offsetHeight;
  bar.style.animation = '';

  try {
    // 1. Get document text
    const essayText = await getDocumentText();
    if (!essayText.trim()) {
      throw new Error('Document is empty.');
    }

    // 2. Call Claude API
    const feedback = await callClaude(essayText);

    // 3. Apply tracked changes + comments
    const result = await applyFeedback(feedback);

    btn.disabled = false;
    btn.textContent = '✨ Review This Essay';
    prog.style.display = 'none';
    showStatus('success', '✅ ' + result.comments + ' comment(s) and ' + result.changes + ' tracked change(s) added.');

  } catch(e) {
    btn.disabled = false;
    btn.textContent = '✨ Review This Essay';
    prog.style.display = 'none';
    showStatus('error', '❌ ' + e.message);
  }
}

// ── Get Document Text ──────────────────────────────────────

function getDocumentText() {
  return Word.run(function(context) {
    const body = context.document.body;
    body.load('text');
    return context.sync().then(function() {
      return body.text;
    });
  });
}

// ── Claude API ─────────────────────────────────────────────

async function callClaude(essayText) {
  const focusAreas = getFocusAreas();
  const prompt = buildPrompt(essayText, focusAreas, intensity, styleExamples);

  const response = await fetch('https://sprightly-cannoli-4b4558.netlify.app/.netlify/functions/proxy', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2500,
      messages: [{ role: 'user', content: prompt }]
    })
  });

  if (!response.ok) {
    const err = await response.json();
    if (response.status === 401) throw new Error('Invalid API key. Please check your key in settings.');
    if (response.status === 429) throw new Error('Too many requests. Please wait a moment and try again.');
    throw new Error(err.error?.message || 'API error ' + response.status);
  }

  const data = await response.json();
  return data.content[0].text;
}

// ── Prompt Builder ─────────────────────────────────────────

function buildPrompt(essayText, focusAreas, intensity, styleExamples) {
  const focusText = focusAreas.length > 0
    ? 'Pay particular attention to: ' + focusAreas.join(', ') + '.'
    : 'Provide well-rounded feedback across all dimensions.';

  const intensityGuide = {
    light: 'Leave 3-5 comments. Be encouraging. Flag only the most important issues.',
    balanced: 'Leave 6-8 comments. Mix genuine praise with specific, actionable critique.',
    thorough: 'Leave 10-14 detailed comments. Be comprehensive — cover argument, evidence, structure, style, and mechanics.'
  }[intensity] || 'Leave 6-8 balanced comments.';

  let styleSection = '';
  if (styleExamples && styleExamples.trim()) {
    styleSection = '\nHere are real examples of how this teacher comments — mimic their tone, vocabulary and directness closely:\n' + styleExamples + '\n\n';
  }

  return 'You are a teacher reviewing a student essay.' + styleSection + '\n'
    + 'Return a JSON object with TWO arrays:\n\n'
    + '1. "comments" — higher-level feedback on argument, structure, evidence, style\n'
    + '2. "suggestions" — precise inline corrections for spelling, punctuation, grammar only\n\n'
    + 'Rules for comments:\n'
    + '- ' + intensityGuide + '\n'
    + '- ' + focusText + '\n'
    + '- Write in first person as the teacher\n'
    + '- Be specific and actionable\n\n'
    + 'Rules for suggestions:\n'
    + '- "find" must be the EXACT verbatim text from the essay\n'
    + '- "replace" is the corrected version\n'
    + '- Only clear-cut errors: spelling, missing punctuation, wrong quote style\n'
    + '- Single word or short phrase only\n'
    + '- Maximum 8 suggestions\n\n'
    + 'Return ONLY valid JSON, no markdown:\n'
    + '{\n'
    + '  "comments": [\n'
    + '    {"quote": "short phrase from essay to anchor comment", "comment": "Feedback here."}\n'
    + '  ],\n'
    + '  "suggestions": [\n'
    + '    {"find": "becuase", "replace": "because"}\n'
    + '  ]\n'
    + '}\n\n'
    + 'Essay:\n\n' + essayText;
}

// ── Parse Response ─────────────────────────────────────────

function parseFeedback(raw) {
  const cleaned = raw.replace(/```json/g, '').replace(/```/g, '').trim();
  try {
    const parsed = JSON.parse(cleaned);
    return {
      comments: (parsed.comments || []).filter(function(x) { return x.comment; }),
      suggestions: (parsed.suggestions || []).filter(function(x) { return x.find && x.replace && x.find !== x.replace; })
    };
  } catch(e) {
    throw new Error('Could not parse AI response. Please try again.');
  }
}

// ── Apply Feedback to Word Document ───────────────────────

async function applyFeedback(rawFeedback) {
  const feedback = parseFeedback(rawFeedback);
  let commentsAdded = 0;
  let changesAdded = 0;

  // Apply tracked changes (suggestions) first
  if (feedback.suggestions && feedback.suggestions.length > 0) {
    changesAdded = await applyTrackedChanges(feedback.suggestions);
  }

  // Then add margin comments
  if (feedback.comments && feedback.comments.length > 0) {
    commentsAdded = await addMarginComments(feedback.comments);
  }

  //return { comments: commentsAdded, changes: changesAdded };
  return {  };
}

// ── Tracked Changes via Office.js ──────────────────────────

async function applyTrackedChanges(suggestions) {
  let applied = 0;

  try {
    await Word.run(async function(context) {
      // Step 1: Enable track changes and sync FIRST
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();

      // Step 2: Apply all replacements in the same context
      for (const s of suggestions) {
        try {
          const results = context.document.body.search(s.find, {
            matchCase: true,
            matchWholeWord: false
          });
          results.load('items');
          await context.sync();

          if (results.items.length > 0) {
            results.items.forEach(function(range) {
              range.insertText(s.replace, Word.InsertLocation.replace);
            });
            await context.sync();
            applied += results.items.length;
          }
        } catch(e) {
          console.log('Tracked change failed for:', s.find, e.message);
        }
      }

      // Step 3: Turn off track changes in the same context
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });
  } catch(e) {
    console.log('Error in applyTrackedChanges:', e.message);
  }

  return applied;
}

// ── Margin Comments via Office.js ──────────────────────────

async function addMarginComments(comments) {
  let added = 0;

  for (const item of comments) {
    try {
      await Word.run(async function(context) {
        let range;

        if (item.quote && item.quote.trim()) {
          // Try to find and anchor to the quoted text
          const results = context.document.body.search(item.quote, {
            matchCase: false,
            matchWholeWord: false
          });
          results.load('items');
          await context.sync();

          if (results.items.length > 0) {
            range = results.items[0];
          }
        }

        // Fallback to end of document if quote not found
        if (!range) {
          range = context.document.body.getRange(Word.RangeLocation.end);
        }

        // Add comment
        range.insertComment(item.comment);
        await context.sync();
        added++;
      });
    } catch(e) {
      console.log('Comment failed:', e.message);
    }
  }

  return added;
}
