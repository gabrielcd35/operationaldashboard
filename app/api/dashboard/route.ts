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