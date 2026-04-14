export async function GET() {
  const url = process.env.APPS_SCRIPT_URL;

  if (!url) {
    return Response.json(
      { error: 'Missing APPS_SCRIPT_URL' },
      { status: 500 }
    );
  }

  try {
    const res = await fetch(url, {
      cache: 'no-store',
      redirect: 'follow',
    });

    if (!res.ok) {
      return Response.json(
        { error: `Failed to fetch Apps Script data: ${res.status}` },
        { status: 500 }
      );
    }

    const text = await res.text();

    if (!text) {
      return Response.json(
        { error: 'Empty response from Apps Script' },
        { status: 500 }
      );
    }

    const json = JSON.parse(text);

    // Normalize date strings in partsRows from "YYYY-MM-DD HH:MM:SS" → "YYYY-MM-DDTHH:MM:SS"
    if (Array.isArray(json.partsRows)) {
      const dateKeys = ['Received At', 'Ordered At', 'Checked Out At', 'Returned At'];
      const spaceDate = /^(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})$/;
      json.partsRows = json.partsRows.map((row: Record<string, unknown>) => {
        const normalized = { ...row };
        for (const key of dateKeys) {
          const val = normalized[key];
          if (typeof val === 'string' && spaceDate.test(val)) {
            normalized[key] = val.replace(' ', 'T');
          }
        }
        return normalized;
      });
    }

    return Response.json(json);
  } catch (error) {
    return Response.json(
      {
        error: error instanceof Error ? error.message : 'Unknown error',
      },
      { status: 500 }
    );
  }
}