export default function mathRandomString(): string {
  return Math.random().toString(36).slice(2);
}
