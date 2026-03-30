export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(),
  });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json();
    const apiKey = context.env.ANTHROPIC_API_KEY || body.apiKey;

    if (!apiKey) {
      return json({ success: false, error: 'Missing Anthropic API key' }, 400);
    }

    if (!body.imageBase64 || !body.mediaType) {
      return json({ success: false, error: 'Missing image payload' }, 400);
    }

    const prompt = `You are extracting invoice line items from a landscaping invoice or handwritten estimate.
Return JSON only with this exact shape:
{
  "description": string,
  "property": string,
  "dateSent": "YYYY-MM-DD" | "",
  "dueDate": "YYYY-MM-DD" | "",
  "invoiceNumber": string,
  "items": [{"name": string, "amount": number}]
}
Rules:
- Extract each service line separately.
- Convert money strings to numbers only.
- If a field is missing, return an empty string.
- Do not include markdown fences.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: body.model || 'claude-sonnet-4-6',
        max_tokens: 1400,
        messages: [
          {
            role: 'user',
            content: [
              {
                type: 'image',
                source: {
                  type: 'base64',
                  media_type: body.mediaType,
                  data: body.imageBase64,
                },
              },
              { type: 'text', text: prompt },
            ],
          },
        ],
      }),
    });

    const payload = await response.json();

    if (!response.ok) {
      return json(
        {
          success: false,
          error: payload?.error?.message || 'Anthropic request failed',
          raw: payload,
        },
        response.status,
      );
    }

    const text = (payload.content || [])
      .map((block) => block.text || '')
      .join('\n')
      .trim();

    const data = safeJson(text);

    return json({ success: true, data, rawText: text }, 200);
  } catch (error) {
    return json({ success: false, error: error.message }, 500);
  }
}

function safeJson(text) {
  if (!text) {
    return {
      description: '',
      property: '',
      dateSent: '',
      dueDate: '',
      invoiceNumber: '',
      items: [],
    };
  }

  const cleaned = text
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```$/i, '')
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch (err) {
    const start = cleaned.indexOf('{');
    const end = cleaned.lastIndexOf('}');
    if (start !== -1 && end !== -1) {
      return JSON.parse(cleaned.slice(start, end + 1));
    }
    throw err;
  }
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json',
  };
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: corsHeaders(),
  });
}
