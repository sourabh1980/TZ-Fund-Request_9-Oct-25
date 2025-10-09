/**
 * ChatEngine namespace
 *
 * This file centralizes chat-related entry points under `Chat.*` without
 * duplicating existing logic. It maps the already-defined global chat
 * functions (from code.gs) into a single object for clearer usage.
 *
 * Usage (server):
 *   Chat.processChatQuery(message)
 *
 * Frontend calls to google.script.run.processChatQuery continue to work
 * unchanged, because the original global function still exists. You can
 * later migrate to only expose Chat.processChatQuery and slim code.gs.
 */
// --- LLM bridges (OpenRouter + DeepSeek) ---
function _chatGetProp_(k){
  try { return PropertiesService.getScriptProperties().getProperty(k) || ''; } catch(e){ return ''; }
}

function askOpenRouter(messages, modelOverride){
  var key = _chatGetProp_('OPENROUTER_API_KEY');
  if(!key) throw new Error('OPENROUTER_API_KEY missing in Script Properties');
  // modelOverride can be a string (model) or an object { model, temperature }
  var defaultModel = _chatGetProp_('OPENROUTER_MODEL') || 'meta-llama/llama-4-maverick:free';
  var model = defaultModel;
  var temperature = 0.2;
  if (modelOverride && typeof modelOverride === 'object') {
    model = modelOverride.model || defaultModel;
    if (typeof modelOverride.temperature === 'number') temperature = modelOverride.temperature;
  } else if (typeof modelOverride === 'string' && modelOverride) {
    model = modelOverride;
  }
  var url = 'https://openrouter.ai/api/v1/chat/completions';
  var payload = {
    model: model,
    messages: messages,
    temperature: temperature,
    max_tokens: 900
  };
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + key,
      'HTTP-Referer': _chatGetProp_('OPENROUTER_REFERER') || 'https://example.com',
      'X-Title': _chatGetProp_('OPENROUTER_TITLE') || 'Fund Request Assistant'
    }
  });
  var code = resp.getResponseCode();
  var dataText = resp.getContentText() || '';
  var data = {};
  try{ data = JSON.parse(dataText); }catch(_){ }
  var out = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if(!out){
    throw new Error('OpenRouter error ' + code + ': ' + (dataText || 'no content'));
  }
  return out;
}

/**
 * DeepSeek direct API (OpenAI-compatible schema)
 * Requires Script Property: DEEPSEEK_API_KEY
 * Optional: DEEPSEEK_MODEL (default: deepseek-chat)
 */
function askDeepSeek(messages, modelOverride){
  var key = _chatGetProp_('DEEPSEEK_API_KEY');
  if(!key) throw new Error('DEEPSEEK_API_KEY missing in Script Properties');
  // modelOverride can be a string (model) or an object { model, temperature }
  var defaultModel = _chatGetProp_('DEEPSEEK_MODEL') || 'deepseek-chat';
  var model = defaultModel;
  var temperature = 0.2;
  if (modelOverride && typeof modelOverride === 'object') {
    model = modelOverride.model || defaultModel;
    if (typeof modelOverride.temperature === 'number') temperature = modelOverride.temperature;
  } else if (typeof modelOverride === 'string' && modelOverride) {
    model = modelOverride;
  }
  var url = 'https://api.deepseek.com/v1/chat/completions';
  var payload = {
    model: model,
    messages: messages,
    temperature: temperature,
    max_tokens: 900
  };
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + key
    }
  });
  var code = resp.getResponseCode();
  var dataText = resp.getContentText() || '';
  var data = {};
  try{ data = JSON.parse(dataText); }catch(_){ }
  var out = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if(!out){
    throw new Error('DeepSeek error ' + code + ': ' + (dataText || 'no content'));
  }
  return out;
}

/**
 * DeepSeek embeddings helper.
 * Returns an embedding array (floats) for a single input string.
 * Requires Script Property: DEEPSEEK_API_KEY
 * Optional: DEEPSEEK_EMBEDDING_MODEL
 */
function askDeepSeekEmbeddings(input, modelOverride){
  var key = _chatGetProp_('DEEPSEEK_API_KEY');
  if(!key) throw new Error('DEEPSEEK_API_KEY missing in Script Properties');
  var model = modelOverride || _chatGetProp_('DEEPSEEK_EMBEDDING_MODEL') || 'deepseek-embedding-3.1';
  var url = 'https://api.deepseek.com/v1/embeddings';
  var payload = {
    model: model,
    input: String(input || '')
  };
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key }
  });
  var code = resp.getResponseCode();
  var dataText = resp.getContentText() || '';
  var data = {};
  try{ data = JSON.parse(dataText); }catch(_){ }
  // Expected shape: { data: [ { embedding: [...] } ], ... }
  if (data && data.data && Array.isArray(data.data) && data.data[0] && data.data[0].embedding){
    return data.data[0].embedding;
  }
  throw new Error('DeepSeek embeddings error ' + code + ': ' + (dataText || 'no embedding'));
}

function _llmProvider_(){
  var p = (_chatGetProp_('LLM_PROVIDER') || '').toLowerCase();
  if (p === 'deepseek' || p === 'openrouter') return p;
  // Auto-detect when not explicitly set
  var hasDeepSeek = !!_chatGetProp_('DEEPSEEK_API_KEY');
  var hasOpenRouter = !!_chatGetProp_('OPENROUTER_API_KEY');
  if (hasDeepSeek) return 'deepseek';
  if (hasOpenRouter) return 'openrouter';
  return 'none';
}

function llmAnswer(question, contextText, modelOverride){
  var system = {
    role: 'system',
    content: [
      'You are a helpful analyst for fund requests and vehicle data.',
      'Use ONLY the provided CONTEXT to answer; if insufficient, ask ONE concise clarifying question.',
      'Prefer short, business-friendly phrasing and precise numbers.',
      'If you ask a clarifying question, end the reply with exactly one question.'
    ].join(' ')
  };
  var user = { role: 'user', content: 'QUESTION:\n' + String(question||'') + '\n\nCONTEXT:\n' + String(contextText||'') };
  var provider = _llmProvider_();
  try {
    if (provider === 'deepseek') return askDeepSeek([system, user], modelOverride);
    if (provider === 'openrouter') return askOpenRouter([system, user], modelOverride);
  } catch (e) {
    var errText = String(e||'');
    // Fallback: if OpenRouter hit rate limit and DeepSeek is configured, switch to DeepSeek
    if (/429|rate limit/i.test(errText) && _chatGetProp_('DEEPSEEK_API_KEY')) {
      return askDeepSeek([system, user], modelOverride);
    }
    // Fallback: if DeepSeek failed and OpenRouter is configured, try OpenRouter
    if (_chatGetProp_('OPENROUTER_API_KEY')) {
      return askOpenRouter([system, user], modelOverride);
    }
    throw e;
  }
  throw new Error('LLM not configured. Set LLM_PROVIDER and API key(s).');
}

/**
 * Open mode answer: allows general knowledge; uses CONTEXT when relevant.
 */
function llmAnswerOpen(question, contextText, modelOverride){
  // Play mode persona: warmer, lighter tone. Still safe and professional.
  var system = {
    role: 'system',
    content: [
      'You are an AI assistant in Play mode.',
      'Persona: warm, witty, supportive, lightly playful. Use contractions and natural, conversational phrasing.',
      'Stay professional and respectful; avoid explicit or NSFW content.',
      'Default style: short paragraphs (2–4 sentences), minimal bullets, no headings unless user asks. Prefer approachable language over formal/business tone.',
      'You may add at most one tasteful emoji if it genuinely aids clarity or tone.',
      'When you are uncertain, say so briefly and suggest a next step.'
    ].join(' ')
  };
  var user = { role: 'user', content: 'QUESTION:\n' + String(question||'') + (contextText ? ('\n\nCONTEXT:\n' + String(contextText||'')) : '') };
  var provider = _llmProvider_();
  // Prefer Play-specific model if configured; raise temperature a bit for more variety
  try {
    if (provider === 'deepseek') {
      var dsPlayModel = _chatGetProp_('DEEPSEEK_MODEL_PLAY') || _chatGetProp_('DEEPSEEK_MODEL') || null;
      return askDeepSeek([system, user], (modelOverride || { model: dsPlayModel, temperature: 0.75 }));
    }
    if (provider === 'openrouter') {
      var orPlayModel = _chatGetProp_('OPENROUTER_MODEL_PLAY') || _chatGetProp_('OPENROUTER_MODEL') || null;
      return askOpenRouter([system, user], (modelOverride || { model: orPlayModel, temperature: 0.75 }));
    }
  } catch (e) {
    var errText = String(e||'');
    // Fallback: if OpenRouter rate-limited and DeepSeek available, try DeepSeek
    if (/429|rate limit/i.test(errText) && _chatGetProp_('DEEPSEEK_API_KEY')) {
      var dsPlayModel2 = _chatGetProp_('DEEPSEEK_MODEL_PLAY') || _chatGetProp_('DEEPSEEK_MODEL') || null;
      return askDeepSeek([system, user], (modelOverride || { model: dsPlayModel2, temperature: 0.75 }));
    }
    // Or if DeepSeek failed and OpenRouter available, try OpenRouter
    if (_chatGetProp_('OPENROUTER_API_KEY')) {
      var orPlayModel2 = _chatGetProp_('OPENROUTER_MODEL_PLAY') || _chatGetProp_('OPENROUTER_MODEL') || null;
      return askOpenRouter([system, user], (modelOverride || { model: orPlayModel2, temperature: 0.75 }));
    }
    throw e;
  }
  throw new Error('LLM not configured. Set LLM_PROVIDER and API key(s).');
}

var Chat = (function(){
  function fn(name){ try { return (typeof thisScope[name] === 'function') ? thisScope[name] : null; } catch(_){ return null; } }
  var thisScope = this;
  return {
    processChatQuery: (typeof processChatQuery === 'function') ? processChatQuery : function(msg){ return 'Chat engine not initialized'; },
    analyzeIntentWithLearning: (typeof analyzeIntentWithLearning === 'function') ? analyzeIntentWithLearning : null,
    analyzeIntent: (typeof analyzeIntent === 'function') ? analyzeIntent : null,
    extractEntities: (typeof extractEntities === 'function') ? extractEntities : null,
    handleFollowupResponse: (typeof handleFollowupResponse === 'function') ? handleFollowupResponse : null,
    isFollowupResponse: (typeof isFollowupResponse === 'function') ? isFollowupResponse : null,
    addFeedbackPrompt: (typeof addFeedbackPrompt === 'function') ? addFeedbackPrompt : null,
    handleRatingFeedback: (typeof handleRatingFeedback === 'function') ? handleRatingFeedback : null,
    logConversation: (typeof logConversation === 'function') ? logConversation : null,
    // LLM bridge
    askOpenRouter: askOpenRouter,
    askDeepSeek: askDeepSeek,
    llmAnswer: llmAnswer,
    llmAnswerOpen: llmAnswerOpen
  };
}).call(this);

/** Simple connectivity test to OpenRouter LLM */
function debugOpenRouterPing(){
  var t0 = new Date().getTime();
  try{
    if(!_chatGetProp_('OPENROUTER_API_KEY')){
      return { ok:false, error:'OPENROUTER_API_KEY missing', ms:(new Date().getTime()-t0) };
    }
    var sys = { role:'system', content:'You are a deterministic tester. Reply with exactly: pong' };
    var usr = { role:'user', content:'Say pong' };
    var txt = askOpenRouter([sys, usr], null) || '';
    var ok = String(txt||'').trim().toLowerCase() === 'pong';
    return { ok: ok, model: _chatGetProp_('OPENROUTER_MODEL') || '', response: txt, ms: (new Date().getTime()-t0) };
  }catch(e){
    return { ok:false, error:String(e), ms:(new Date().getTime()-t0) };
  }
}

/** Simple connectivity test to DeepSeek LLM */
function debugDeepSeekPing(){
  var t0 = new Date().getTime();
  try{
    if(!_chatGetProp_('DEEPSEEK_API_KEY')){
      return { ok:false, error:'DEEPSEEK_API_KEY missing', ms:(new Date().getTime()-t0) };
    }
    var sys = { role:'system', content:'You are a deterministic tester. Reply with exactly: pong' };
    var usr = { role:'user', content:'Say pong' };
    var txt = askDeepSeek([sys, usr], null) || '';
    var ok = String(txt||'').trim().toLowerCase() === 'pong';
    return { ok: ok, model: _chatGetProp_('DEEPSEEK_MODEL') || '', response: txt, ms: (new Date().getTime()-t0) };
  }catch(e){
    return { ok:false, error:String(e), ms:(new Date().getTime()-t0) };
  }
}

/**
 * LLM configuration helpers (secure Script Properties setup)
 * Note: Do NOT log secrets; these helpers avoid printing provided keys.
 */
function _setPropSafe_(k, v){
  try {
    PropertiesService.getScriptProperties().setProperty(k, String(v||''));
    return true;
  } catch(e){ return false; }
}

function configureOpenRouter(apiKey, model, title, referer, makeDefault){
  if (!apiKey) throw new Error('Missing apiKey');
  var ok1 = _setPropSafe_('OPENROUTER_API_KEY', apiKey);
  if (model) _setPropSafe_('OPENROUTER_MODEL', model);
  if (title) _setPropSafe_('OPENROUTER_TITLE', title);
  if (referer) _setPropSafe_('OPENROUTER_REFERER', referer);
  if (makeDefault === true) _setPropSafe_('LLM_PROVIDER', 'openrouter');
  return getLLMConfigSummary();
}

function configureDeepSeek(apiKey, model, makeDefault){
  if (!apiKey) throw new Error('Missing apiKey');
  var ok1 = _setPropSafe_('DEEPSEEK_API_KEY', apiKey);
  if (model) _setPropSafe_('DEEPSEEK_MODEL', model);
  if (makeDefault === true) _setPropSafe_('LLM_PROVIDER', 'deepseek');
  return getLLMConfigSummary();
}

function getLLMConfigSummary(){
  var sp = PropertiesService.getScriptProperties();
  var prov = (_chatGetProp_('LLM_PROVIDER')||'').toLowerCase();
  var hasOR = !!_chatGetProp_('OPENROUTER_API_KEY');
  var hasDS = !!_chatGetProp_('DEEPSEEK_API_KEY');
  var modelOR = _chatGetProp_('OPENROUTER_MODEL') || '';
  var modelDS = _chatGetProp_('DEEPSEEK_MODEL') || '';
  var modelORPlay = _chatGetProp_('OPENROUTER_MODEL_PLAY') || '';
  var modelDSPlay = _chatGetProp_('DEEPSEEK_MODEL_PLAY') || '';
  return {
    provider: prov || (hasDS ? 'deepseek' : (hasOR ? 'openrouter' : 'none')),
    openrouter: { keySet: hasOR, model: modelOR, playModel: modelORPlay },
    deepseek: { keySet: hasDS, model: modelDS, playModel: modelDSPlay }
  };
}

function clearLLMKeys(){
  var sp = PropertiesService.getScriptProperties();
  ['OPENROUTER_API_KEY','OPENROUTER_MODEL','OPENROUTER_MODEL_PLAY','OPENROUTER_TITLE','OPENROUTER_REFERER','DEEPSEEK_API_KEY','DEEPSEEK_MODEL','DEEPSEEK_MODEL_PLAY']
    .forEach(function(k){ try{ sp.deleteProperty(k); }catch(_){ } });
  return getLLMConfigSummary();
}

// Convenience: update OpenRouter non-secret settings without changing the API key
function updateOpenRouterSettings(model, referer, title){
  if (model) _setPropSafe_('OPENROUTER_MODEL', model);
  if (referer) _setPropSafe_('OPENROUTER_REFERER', referer);
  if (title) _setPropSafe_('OPENROUTER_TITLE', title);
  return getLLMConfigSummary();
}

// Set a dedicated Play model for DeepSeek (optional)
function updateDeepSeekPlayModel(model){
  if (!model) throw new Error('Missing model');
  _setPropSafe_('DEEPSEEK_MODEL_PLAY', model);
  return getLLMConfigSummary();
}

// Set a dedicated Play model for OpenRouter (optional)
function updateOpenRouterPlayModel(model){
  if (!model) throw new Error('Missing model');
  _setPropSafe_('OPENROUTER_MODEL_PLAY', model);
  return getLLMConfigSummary();
}

// A/B test: compare Work vs Play responses and log both
function debugPlayVsWork(){
  var q = 'Give two friendly tips for thanking a donor.';
  var ctx = '';
  var work = llmAnswer(q, ctx, { temperature: 0.2 });
  var play = llmAnswerOpen(q, ctx, { temperature: 0.9 });
  Logger.log('WORK -> ' + work);
  Logger.log('PLAY -> ' + play);
  return { provider: _llmProvider_(), work: work, play: play };
}

// Convenience direct calls
function forcePlayAnswer(question, context){
  return llmAnswerOpen(question, context || '', { temperature: 0.9 });
}
function forceWorkAnswer(question, context){
  return llmAnswer(question, context || '', { temperature: 0.2 });
}
