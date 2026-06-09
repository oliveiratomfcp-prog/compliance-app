const fetch = require('node-fetch');

module.exports = async function (context, req) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ error: { message: 'OPENAI_API_KEY non configurée sur le serveur.' } }) };
    return;
  }

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify(req.body)
    });

    const data = await response.json();
    context.res = {
      status: response.status,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    };
  } catch (e) {
    context.res = {
      status: 502,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: { message: e.message } })
    };
  }
};
