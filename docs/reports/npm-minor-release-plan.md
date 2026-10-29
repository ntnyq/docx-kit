# npm minor release plan

Research date: 2026-10-29. Target release: `0.7.0`.

This is a proposed runbook. No packages, tags, releases, or npm settings were
changed during research. Publishing requires the maintainer's confirmation.

## Registry inventory

The 40 public workspace packages currently declare `0.6.0`. Registry checks
found 35 existing packages with `latest` pointing to `0.6.0`, and five names
returning HTTP 404:

| New package | Workspace directory |
| --- | --- |
| `@docxkit/errors` | `packages/errors` |
| `@docxkit/markdown` | `packages/markdown` |
| `@docxkit/plugin-checklist` | `packages-plugins/checklist` |
| `@docxkit/plugin-figure` | `packages-plugins/figure` |
| `@docxkit/plugin-metric-card` | `packages-plugins/metric-card` |

The root, documentation, applications, and internal shared package are private.
Recheck the inventory immediately before publishing; public HTTP 404 alone does
not prove that an authenticated account has permission to create a name.

The existing `.github/workflows/release.yml` grants `id-token: write`, runs on
`v*` tags, verifies the release, and publishes recursively. The
[v0.6.0 release run](https://github.com/ntnyq/docx-kit/actions/runs/40638647314)
completed successfully, including its publish step. Existing trust settings
have not been individually inspected during this research.

## Batch trust configuration

Use the official `npm trust` command, available since npm 11.15.0. It accepts
one package per invocation and does not discover pnpm workspaces.

The official bulk workflow is to complete the first browser 2FA challenge and
select the option to skip 2FA for five minutes. Then run subsequent commands
sequentially with `--yes`, leaving two seconds between requests. `--yes` skips
the ordinary confirmation prompt; it does not disable 2FA. See
[npm trust: prerequisites, flags, and bulk usage](https://docs.npmjs.com/cli/v11/commands/npm-trust/).

Run the following only after the five names exist and publishing is approved:

```bash
new_packages=(
  @docxkit/errors
  @docxkit/markdown
  @docxkit/plugin-checklist
  @docxkit/plugin-figure
  @docxkit/plugin-metric-card
)

for package_name in "${new_packages[@]}"; do
  npm trust github "$package_name" \
    --repo ntnyq/docx-kit \
    --file release.yml \
    --allow-publish \
    --yes || break
  sleep 2
done
```

Use an interactive terminal for browser authentication. The workflow has no
GitHub environment, so omit `--env`. Preserve any already matching connection
instead of deleting and recreating it. If a command fails, record the error and
resume from the affected package after checking its state.

## Recommended release sequence

1. Recheck the registry inventory and publication rights. Run
   `pnpm run release:verify`, review changes since `v0.6.0`, and prepare the
   version changes and release notes. Complete builds and packing before opening
   a five-minute authentication window.
2. Prepare real `0.7.0-rc.0` tarballs for the five new packages with `pnpm pack`.
   Use temporary preparation state; keep the final release manifests at
   `0.7.0`. Inspect packed manifests so dependency versions match the intended
   release graph and contain no `workspace:` protocols. These are bootstrap
   prereleases, not installation candidates before their dependencies exist.
3. Authenticate using `npm login --auth-type=web`. Sequentially publish those
   prebuilt tarballs with `npm publish <tarball> --access public --tag bootstrap`.
   Only the five new names receive this prerelease. Existing packages retain
   their current `latest` versions. Use real artifacts, not empty placeholders.
4. Configure the five trusted publishers using the loop above. At the first
   trust challenge, enable the five-minute skip option. Read back the resulting
   connections while the authentication window is active.
5. Restore/recheck all final manifests at `0.7.0`, refresh the lockfile as needed,
   and rerun the release checks against the final state. Commit the reviewed
   version changes, then push the release commit and `v0.7.0` tag.
6. Let the existing Actions workflow publish all 40 stable packages using OIDC.
   Verify every `0.7.0` tarball and `latest` tag, the Actions result, and consumer
   installation. Remove obsolete bootstrap dist-tags after verification; the
   prerelease versions remain in npm history.

The prerelease step ensures that the package exists before configuring trust,
while reserving the stable version for an actual OIDC publish. New trust
connections must complete their first successful publish within two days;
otherwise they expire. The stable Actions publish validates them immediately.
See [trusted publishing and configuration expiry](https://docs.npmjs.com/trusted-publishers/).

Recursive pnpm publishing skips versions already present in the registry. Thus
manually publishing stable `0.7.0` for new packages first would cause CI to skip
them, leaving their new trust configurations unvalidated. Also, recursive
publishing is not an atomic multi-package transaction; keep a publication
summary and inspect partial results before retrying. See
[pnpm publish](https://pnpm.io/cli/publish).

## Authentication budget and fallback

Prefer a temporary interactive npm session so no additional publish token is
needed. Plan for approximately two or three user interactions: login, the
first manual publish, and the first trust configuration. This is an estimate,
not a guaranteed maximum. The official five-minute guidance explicitly covers
the trust loop; reuse across publish and trust has not been live-tested.

Both commands use npm's `otplease` helper, but the CLI does not persist a
cross-command OTP cache. Any skip window is controlled by npm's service. See
the npm 11.19.1 sources for
[authentication](https://github.com/npm/cli/blob/v11.19.1/lib/utils/auth.js),
[publishing](https://github.com/npm/cli/blob/v11.19.1/lib/commands/publish.js), and
[trust configuration](https://github.com/npm/cli/blob/v11.19.1/lib/trust-cmd.js).

If interactive bootstrap publishing repeatedly requests authentication, a
short-lived granular token with write access to the `@docxkit` scope and
Bypass 2FA is a fallback for those five publishes. It cannot configure trust;
that requires an interactive session. Giving a token organization-management
permissions alone does not grant package publication rights. npm currently
allows direct-capable bypass tokens to publish, but plans to remove this
capability in January 2027. Do not build the ongoing release workflow around
one. See [access token permissions and deprecation](https://docs.npmjs.com/about-access-tokens/)
and [registry trust API requirements](https://api-docs.npmjs.com/).

## Research validation

- Queried the public registry for all 40 package names.
- Checked the last successful release workflow and its publish step.
- Confirmed CLI help supports the proposed trust flags.
- Packed all five new packages into temporary local archives without upload.
  Each contained its JavaScript and declaration entry points; their packed
  manifests had converted workspace dependencies to concrete versions.
- Did not run the full release verification suite, change versions, start an
  authentication challenge, create credentials, configure trust, or publish.

The research archives contain the current `0.6.0` development state and are not
release inputs. Regenerate artifacts from the approved release state.
