/**
 * Resume/CV template.
 *
 * Generates a professional resume with personal info,
 * experience, education, and skills sections.
 *
 * @module ai/templates/resume
 */

import type { DocxSchema } from '../../builder/createDocx'
import type { AiTemplate } from '../types'

/** Resume template parameters. */
export interface ResumeParams {
  /** Full name. */
  name: string
  /** Email address. */
  email?: string
  /** Phone number. */
  phone?: string
  /** Skills list. */
  skills?: string[]
  /** Professional summary / objective. */
  summary?: string
  /** Education entries. */
  education?: {
    /** Degree/major. */
    degree?: string
    /** Institution name. */
    institution?: string
    /** Graduation year. */
    year?: string
  }[]
  /** Work experience entries. */
  experience?: {
    /** Company name. */
    company?: string
    /** End date or "Present". */
    endDate?: string
    /** Role description bullets. */
    highlights?: string[]
    /** Role/title. */
    role?: string
    /** Start date. */
    startDate?: string
  }[]
}

const systemPrompt = `You are a professional document generator. Generate a docx-kit JSON schema for a resume/CV document.

The resume should include:
1. Name and contact info header
2. Professional summary
3. Work experience with role, company, dates, and highlights
4. Education with institution, degree, and year
5. Skills section

Use the following docx-kit node types:
- { type: "heading", level: N, text: "..." } for section headers
- { type: "paragraph", text: "..." } for text content
- { type: "bulletList", items: [...] } for bulleted lists
- { type: "plugin", name: "propertyTable", options: { ... } } for contact info

Maintain a professional tone appropriate for the resume type.
`

const schema = {
  description: 'Professional resume/CV with experience, education, and skills',
  title: 'ResumeParams',
  type: 'object',
  properties: {
    email: { description: 'Email address', type: 'string' },
    name: { description: 'Full name', required: true, type: 'string' },
    phone: { description: 'Phone number', type: 'string' },
    education: {
      description: 'Education entries',
      type: 'array',
      items: {
        type: 'object',
        properties: {
          degree: { description: 'Degree/major', type: 'string' },
          institution: { description: 'Institution name', type: 'string' },
          year: { description: 'Graduation year', type: 'string' },
        },
      },
    },
    experience: {
      description: 'Work experience entries',
      type: 'array',
      items: {
        type: 'object',
        properties: {
          company: { description: 'Company name', type: 'string' },
          endDate: { description: 'End date or "Present"', type: 'string' },
          role: { description: 'Role/title', type: 'string' },
          startDate: { description: 'Start date', type: 'string' },
          highlights: {
            description: 'Role description bullets',
            items: { type: 'string' },
            type: 'array',
          },
        },
      },
    },
    skills: {
      description: 'Skills list',
      items: { type: 'string' },
      type: 'array',
    },
    summary: {
      description: 'Professional summary / objective',
      type: 'string',
    },
  },
}

function generate(params: ResumeParams): DocxSchema {
  const content: any[] = [{ level: 1, text: params.name, type: 'heading' }]

  // Name header

  // Contact info

  const contactItems: any[] = []
  if (params.email) {
    contactItems.push({ key: 'Email', value: params.email })
  }
  if (params.phone) {
    contactItems.push({ key: 'Phone', value: params.phone })
  }

  if (contactItems.length > 0) {
    content.push({
      name: 'propertyTable',
      options: { items: contactItems },
      type: 'plugin',
    })
  }

  // Summary
  if (params.summary) {
    content.push(
      { level: 2, text: 'Summary', type: 'heading' },
      { text: params.summary, type: 'paragraph' },
    )
  }

  // Experience
  if (params.experience && params.experience.length > 0) {
    content.push({ level: 2, text: 'Experience', type: 'heading' })
    for (const exp of params.experience) {
      const title = exp.role ?? 'Position'
      const company = exp.company ? ` — ${exp.company}` : ''
      const dates =
        exp.startDate && exp.endDate
          ? ` (${exp.startDate} – ${exp.endDate})`
          : ''
      content.push({
        level: 3,
        text: `${title}${company}${dates}`,
        type: 'heading',
      })
      if (exp.highlights) {
        content.push({ items: exp.highlights, type: 'bulletList' })
      }
    }
  }

  // Education
  if (params.education && params.education.length > 0) {
    content.push({ level: 2, text: 'Education', type: 'heading' })
    for (const edu of params.education) {
      const text = `${edu.degree ?? 'Degree'}, ${edu.institution ?? 'University'}${edu.year ? ` (${edu.year})` : ''}`
      content.push({ text, type: 'paragraph' })
    }
  }

  // Skills
  if (params.skills && params.skills.length > 0) {
    content.push(
      { level: 2, text: 'Skills', type: 'heading' },
      { text: params.skills.join(' • '), type: 'paragraph' },
    )
  }

  return { content } as any as DocxSchema
}

export const resumeTemplate: AiTemplate<ResumeParams> = {
  description: 'Professional resume/CV with experience, education, and skills',
  generate,
  name: 'resume',
  schema: schema as any,
  systemPrompt,
}
