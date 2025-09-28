# Security Policy

## Supported Versions
| Version | Supported |
| ------- | --------- |
| 0.x     | Best-effort security fixes |

## Reporting a Vulnerability
Please open a private security advisory or email: security@example.com (replace with a real contact).

Provide:
- Affected version or commit hash
- Reproduction JSON payload (if applicable)
- Expected vs actual behaviour
- Impact assessment (data leak, DoS, etc.)

We aim to respond within 5 working days. Public disclosure should only happen after a coordinated fix window.

## Hardening Recommendations
- Run generation in an isolated worker/container.
- Disable remote image download in untrusted multi-tenant scenarios.
- Enforce max slides and max elements per slide.
- Use only whitelisted domains for `url:` images.

## Planned Improvements
- Optional sandbox for image fetching
- Configurable memory / time guards
