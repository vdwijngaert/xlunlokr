export function unlockedFilename(name: string): string {
  const dot = name.lastIndexOf(".");
  if (dot <= 0) return `${name}-unlocked`;
  return `${name.slice(0, dot)}-unlocked${name.slice(dot)}`;
}
