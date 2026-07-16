/** Convert a plugin name into a safe camelCase JavaScript identifier. */
export function toIdentifier(pluginName: string): string {
  const identifier = pluginName
    .replace(/^[^a-z_$]+/i, '')
    .replace(/[^\w$]+([\w$])?/g, (_, next: string | undefined) =>
      next ? next.toUpperCase() : '',
    )

  return identifier || 'plugin'
}

/** Capitalize an identifier for use in exported type names. */
export function toPascalCase(pluginName: string): string {
  const identifier = toIdentifier(pluginName)
  return identifier.charAt(0).toUpperCase() + identifier.slice(1)
}
