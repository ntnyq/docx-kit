export interface DocumentOptions {
  creator?: string
  description?: string
  title?: string
}

export interface TitleOptions {
  alignment?: 'center' | 'left' | 'right'
  color?: string
  font?: string
  fontSize?: number
  level?: number
  spacing?: {
    after?: number
    before?: number
  }
}
