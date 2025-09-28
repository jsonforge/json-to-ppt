# Commercial Strategy (MIT + Hosted Service)

## Vision
Turn `json-to-ppt` into the default declarative slide generation layer for internal automation, LLM agents and reporting pipelines, while monetising via hosted API, premium templates, collaboration & governance features.

## Monetisation Pillars
| Pillar | Free (MIT Core) | Paid / Hosted Upsell |
| ------ | ---------------- | -------------------- |
| Core Generation | Local library | Scalable API (autoscaling, queueing, retries) |
| Templates | Basic examples | Curated premium packs, vertical (Finance, Edu, Marketing) |
| Collaboration | N/A | Team workspaces, version history, role based access |
| Compliance / Audit | N/A | Activity logs, watermark, deterministic hash registry |
| Performance | Depends on user infra | Optimised rendering cluster, priority lanes |
| Extensions | Manual merge | One‑click install from template marketplace |
| Support | Community issues | SLA, private escalation channel |

## Phased Roadmap
### Phase 0 – Foundation (Current)
- Solidify schema stability & backward compatibility contract
- Add semantic versioning & CHANGELOG (planned)

### Phase 1 – Hosted MVP
- REST API: `/v1/render` (sync small, async large) + `/v1/jobs/:id`
- Auth: API Key header `X-API-Key`
- Rate limiting (basic fixed window) + usage meters (slides rendered)
- Billing metric: slide count + generated file size tiers

### Phase 2 – Templates Marketplace
- Template definition spec (metadata + preview + JSON body)
- Signing / integrity hash to avoid tampering
- Tagging: industry / layout / style
- Paid bundles with license key delivery

### Phase 3 – Teams & Governance
- Org + Members + Roles (admin/editor/viewer)
- Shared asset library (images / brand palette)
- Activity log & regenerate history (hash compare)

### Phase 4 – Advanced Value Adds
- Diff visualisation between two JSON payloads
- Regression snapshot service (binary pptx hash store)
- AI Assist: natural language → partial structured JSON patches

## Pricing Sketch (Indicative)
| Plan | Monthly Price | Included Slides | Overage | Extras |
| ---- | ------------- | -------------- | ------- | ------ |
| Dev | $0 | 500 | Limited | Rate limit, single key |
| Pro | $29 | 10k | $2 / 1k | 3 keys, priority queue |
| Team | $99 | 50k | $1.5 / 1k | Roles, history, template packs |
| Enterprise | Custom | Unlimited | Custom | SSO, audit, on-prem option |

## Key Moats (Non-License)
- Strong schema + determinism contract (hard to replicate edge-cases quickly)
- Marketplace network effects (template creators)
- Operational excellence (fast, reliable rendering cluster)
- Brand trust & security posture (SOC2 later)

## KPIs
- Time to first slide (<3 min signup → PPT)
- Conversion Free→Paid (% unique users who exceed free quota)
- Template attach rate (% jobs using marketplace templates)
- Churn (monthly logos & revenue)

## Risks & Mitigations
| Risk | Impact | Mitigation |
| ---- | ------ | ---------- |
| Competitor forks core | Medium | Speed of execution + templates + hosted reliability |
| API abuse / overrun | High | Per-key quotas + anomaly detection |
| Large PPT resource exhaustion | Medium | Async job offloading + preflight validation |
| Template piracy | Medium | Hash signing + watermark preview |

## Immediate Next Steps
1. Add CHANGELOG & semantic versioning discipline.
2. Design API spec (see `API_SPEC.md`).
3. Implement minimal hosted prototype (FastAPI / queue worker).
4. Metrics instrumentation (Prometheus counters / simple db aggregation).
5. Launch closed beta with manual onboarding.
