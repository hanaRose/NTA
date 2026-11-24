export function readOnboardingTaskContentUrl(value: unknown, baseUrl: string): string {
  const rawUrl: unknown = typeof value === 'object' && value !== null && 'Url' in value
    ? value.Url
    : value;
  if (typeof rawUrl !== 'string' || !rawUrl.trim()) {
    throw new Error('A task item has no valid ContentUrl value.');
  }

  let parsedUrl: URL;
  try {
    parsedUrl = new URL(rawUrl.trim(), baseUrl);
  } catch {
    throw new Error('A task item has an invalid ContentUrl value.');
  }

  if (parsedUrl.protocol !== 'http:' && parsedUrl.protocol !== 'https:') {
    throw new Error('Task ContentUrl values must use HTTP or HTTPS.');
  }

  return parsedUrl.toString();
}
