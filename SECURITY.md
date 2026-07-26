# Security policy

## Supported versions

Security fixes are provided for the latest published docx-kit release. Users
should reproduce a suspected issue on the latest version before reporting it
when that can be done safely. Older releases may be asked to upgrade rather
than receive a backport.

## Reporting a vulnerability

Please do not disclose suspected vulnerabilities in public issues, pull
requests, discussions, or social media.

Use [GitHub private vulnerability
reporting](https://github.com/ntnyq/docx-kit/security/advisories/new). Include:

- affected package and version;
- Node.js/browser and operating-system versions;
- a minimal reproduction or proof of concept;
- the security impact and required attacker capabilities;
- whether plugin loading, remote URLs, file access, generated DOCX content, or
  the browser preview is involved;
- any known workaround.

Remove secrets and personal document content from reproductions. If an
attachment is necessary, use synthetic data.

The maintainer will acknowledge the report, assess affected versions, and
coordinate remediation and disclosure through the private advisory. Response
times depend on severity and maintainer availability.

## Security boundaries

Third-party plugins execute JavaScript with the permissions of the host
application unless the host supplies additional isolation. A valid plugin
manifest and loader security policy are validation and authorization layers;
they are not a sandbox. Review plugin source and provenance before loading it,
especially from local, npm, or URL sources.

Generated DOCX files and uploaded documents should be treated as untrusted
input when they cross a trust boundary. Keep LibreOffice, Microsoft Word,
browsers, Node.js, and transitive dependencies up to date.
