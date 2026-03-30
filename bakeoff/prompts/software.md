# Prompt 2: Software — Monolith to Microservices Migration

Create a 6-slide technical presentation for "Project Chimera: Monolith Decomposition Plan" — presented to VP Engineering and 20 backend engineers at a mid-stage B2B SaaS company. Dense, data-driven slides. Every slide should have a table, diagram, or metrics. No filler.

---

## Current State: The Monolith

**System:** "Atlas" — a Django monolith (Python 3.11, PostgreSQL 15, Redis 7, Celery workers). 340K lines of Python, 2,800 database tables, 14 Django apps in a single repo.

**Team:** 22 backend engineers across 4 squads, 6 frontend engineers, 3 SREs, 2 DBAs. All commit to `main` in a single repo.

### Pain Points (measured, not anecdotal)

| Metric | Current | Industry P50 (DORA) | Gap |
|--------|---------|---------------------|-----|
| Deploy frequency | 1.8/week (every 3.9 days) | 1/day — 1/week | At bottom of "medium" |
| Lead time for changes | 12.3 days (commit → prod) | 1 day — 1 week | 2x over "medium" ceiling |
| Change failure rate | 18.2% (hotfixes / deploys) | 0-15% | Above "medium" threshold |
| Mean time to recover | 4.2 hours | < 1 hour — < 1 day | Functional but slow |
| Build time (CI) | 47 minutes (full test suite) | — | Developers context-switch |
| Merge conflicts/week | 3.2 avg (peaks to 8 during releases) | — | Cross-team coordination tax |
| Deployment rollback rate | 1 in 5.5 deploys | — | Low confidence in releases |
| Test flakiness | 6.8% of CI runs fail on unrelated tests | — | "Retry and pray" culture |

**Cost estimate:** Engineering time lost to monolith friction: ~1,400 engineer-hours/quarter (based on deployment delays, conflict resolution, rollback recovery, and flaky test investigation). At $95/hr loaded cost = **~$133K/quarter wasted**.

### Architecture Diagram (Current)

```
                    ┌──────────────────────────────────┐
                    │            NGINX                  │
                    │     (reverse proxy + SSL)         │
                    └──────────────┬───────────────────┘
                                   │
                    ┌──────────────▼───────────────────┐
                    │         Django Monolith           │
                    │                                    │
                    │  ┌──────┐ ┌──────┐ ┌──────────┐  │
                    │  │Users │ │Orders│ │Payments  │  │
                    │  ├──────┤ ├──────┤ ├──────────┤  │
                    │  │Search│ │Notif.│ │Analytics │  │
                    │  ├──────┤ ├──────┤ ├──────────┤  │
                    │  │Billing│ │Auth  │ │Reporting │  │
                    │  └──────┘ └──────┘ └──────────┘  │
                    │                                    │
                    │  14 Django apps, shared models     │
                    │  340K LOC, 2800 tables              │
                    └──────────────┬───────────────────┘
                                   │
                    ┌──────────────▼───────────────────┐
                    │      PostgreSQL 15 (single)       │
                    │      + Redis 7 (cache/queue)      │
                    │      + Celery (background jobs)   │
                    └──────────────────────────────────┘
```

---

## Target Architecture

**Pattern:** Strangler Fig — incrementally extract services behind an API gateway. No big bang rewrite.

### Service Boundaries (domain-driven)

| Service | Owner Squad | DB | Key Entities | API Style | Priority |
|---------|------------|-----|-------------|-----------|----------|
| **User Service** | Identity | PostgreSQL (isolated) | Users, Orgs, Roles, Permissions | REST + gRPC | P0 — extract first |
| **Order Service** | Commerce | PostgreSQL (isolated) | Orders, LineItems, Carts, Pricing | REST + events | P0 — highest coupling |
| **Payment Service** | Commerce | PostgreSQL (isolated) + Stripe | Payments, Invoices, Refunds | REST (sync for checkout) | P1 — after orders |
| **Search Service** | Discovery | Elasticsearch 8 | Product index, query logs | REST | P1 — already loosely coupled |
| **Notification Service** | Platform | PostgreSQL + SQS | Templates, DeliveryLog, Preferences | Async (event-driven) | P2 — low risk |
| **Analytics Service** | Data | ClickHouse | Events, Metrics, Funnels | gRPC (internal only) | P2 — read-only workload |

### Target Diagram

```
  Clients (Web, Mobile, API)
           │
    ┌──────▼──────┐
    │ API Gateway  │  Kong / AWS ALB
    │ (routing +   │  Rate limiting, auth token validation
    │  rate limit) │  Circuit breaker per service
    └──────┬──────┘
           │
    ┌──────▼──────────────────────────────────┐
    │            Service Mesh (Istio)          │
    │                                          │
    │  ┌────────┐ ┌────────┐ ┌─────────────┐  │
    │  │ User   │ │ Order  │ │  Payment    │  │
    │  │Service │ │Service │ │  Service    │  │
    │  └───┬────┘ └───┬────┘ └──────┬──────┘  │
    │      │          │             │          │
    │  ┌───▼──────────▼─────────────▼───────┐  │
    │  │        Event Bus (Kafka)           │  │
    │  │  Topics: user.*, order.*, payment.*│  │
    │  └───┬──────────┬─────────────┬───────┘  │
    │      │          │             │          │
    │  ┌───▼────┐ ┌───▼────┐ ┌─────▼───────┐  │
    │  │Search  │ │Notif.  │ │ Analytics   │  │
    │  │Service │ │Service │ │ Service     │  │
    │  └────────┘ └────────┘ └─────────────┘  │
    └──────────────────────────────────────────┘
           │
    ┌──────▼──────────────────────────────────┐
    │        Isolated Data Stores              │
    │  PostgreSQL × 4  │  Elasticsearch  │     │
    │  ClickHouse      │  Redis (per-svc)│     │
    └──────────────────────────────────────────┘
```

---

## Migration Plan — 6-Month Phased Rollout

### Phase 1: Foundation (Month 1-2)

| Task | Owner | Deliverable | Risk |
|------|-------|-------------|------|
| Deploy API Gateway (Kong) in front of monolith | SRE | All traffic routes through gateway; monolith is just another upstream | Gateway becomes SPOF — need HA config |
| Instrument monolith with OpenTelemetry | SRE + all squads | Distributed tracing on all HTTP handlers and Celery tasks | Performance overhead ~2-3% |
| Extract User Service | Identity squad | User CRUD, auth, permissions as standalone service behind gateway | Session migration — must support both old cookie + new JWT during transition |
| Set up Kafka cluster (3 brokers) | SRE | Event bus for async communication; `user.created`, `user.updated` events | Operational complexity — team has no Kafka experience |
| Database migration tooling | DBA | Scripts to split shared tables into per-service schemas; dual-write during transition | Data consistency during migration window |

### Phase 2: Commerce Core (Month 3-4)

| Task | Owner | Deliverable | Risk |
|------|-------|-------------|------|
| Extract Order Service | Commerce squad | Orders, carts, pricing as standalone; consumes user events via Kafka | Highest-coupling extraction — 47 cross-module imports to untangle |
| Extract Payment Service | Commerce squad | Stripe integration, invoicing, refunds as standalone | PCI compliance scope changes — security review required |
| Implement saga pattern for checkout | Commerce + Identity | Distributed transaction: create order → reserve inventory → charge payment | Failure modes are complex; need compensation logic for each step |
| Search Service extraction | Discovery squad | Elasticsearch queries as standalone service; monolith writes to Kafka, search consumes | Index rebuild takes 4 hours — need zero-downtime reindex strategy |

### Phase 3: Platform & Optimization (Month 5-6)

| Task | Owner | Deliverable | Risk |
|------|-------|-------------|------|
| Notification Service extraction | Platform squad | Email/SMS/push via SQS; template management | Low risk — already loosely coupled |
| Analytics Service extraction | Data squad | ClickHouse ingestion from Kafka; dashboards migrate to new endpoints | Historical data migration — 18 months of event data |
| Decommission monolith modules | All squads | Remove extracted code from monolith; monolith becomes thin orchestration layer | Residual coupling — some shared utilities may still reference old code |
| Performance tuning + chaos engineering | SRE | Latency targets met, circuit breakers tested, runbook for each failure mode | Unknown unknowns — distributed systems fail in surprising ways |

---

## Risk Matrix

| Risk | Likelihood | Impact | Severity | Mitigation |
|------|-----------|--------|----------|------------|
| Distributed transaction failures (checkout saga) | HIGH | HIGH | **CRITICAL** | Implement compensation/rollback for each saga step; extensive integration testing; feature flag to fall back to monolith checkout |
| Kafka cluster instability (team inexperience) | MEDIUM | HIGH | **HIGH** | Hire/contract Kafka specialist for first 3 months; use managed Kafka (Confluent Cloud) instead of self-hosted; start with 3 low-throughput topics |
| Data inconsistency during dual-write migration | HIGH | MEDIUM | **HIGH** | CDC (Change Data Capture) with Debezium instead of application-level dual-write; reconciliation jobs running hourly; rollback procedure documented |
| Service-to-service latency exceeds budget | MEDIUM | MEDIUM | **MEDIUM** | P99 latency budget of 200ms per service call; circuit breakers (Istio) with 500ms timeout; async where possible (events > sync calls) |
| Team cognitive overload (too many new tools) | MEDIUM | MEDIUM | **MEDIUM** | Phase tool adoption — Kafka in M1, Istio in M3, ClickHouse in M5; pair programming with SRE on infrastructure tasks; weekly architecture office hours |
| PCI scope expansion for Payment Service | LOW | HIGH | **MEDIUM** | Engage security team in M2 before extraction begins; use Stripe's hosted checkout to minimize PCI surface; document compliance boundaries |

---

## Monitoring & Observability Stack

| Layer | Tool | What It Covers | Alert Threshold |
|-------|------|---------------|-----------------|
| **Metrics** | Prometheus + Grafana | Request rate, error rate, duration (RED); saturation (CPU, memory, connections) | Error rate > 1% for 5 min; P99 > 500ms for 10 min |
| **Tracing** | Jaeger + OpenTelemetry | End-to-end request flow across services; latency breakdown per hop | Trace duration > 2s; span error rate > 5% |
| **Logging** | ELK Stack (Elasticsearch, Logstash, Kibana) | Structured JSON logs; correlation IDs across services | Error log rate > 50/min per service |
| **Health checks** | Kubernetes liveness + readiness probes | Service availability; dependency health (DB, Redis, Kafka) | Any probe failure → restart; 3 consecutive → page SRE |
| **Synthetic monitoring** | Datadog Synthetics | Critical user flows: login, search, checkout, payment | Any failure → immediate page |
| **Alerting** | PagerDuty | Tiered escalation: P1 (page) → P2 (Slack) → P3 (next business day) | See thresholds above |

**Key dashboards (Grafana):**
- Service Overview: request rate, error %, P50/P95/P99 latency per service
- Kafka Health: consumer lag per topic, broker disk usage, replication status
- Checkout Saga: success rate, avg duration, failure breakdown by step
- Deployment Tracker: last deploy time per service, canary status, rollback count

---

## Success Metrics — Before/After Targets

| Metric | Current (Monolith) | 6-Month Target | 12-Month Target | Measurement |
|--------|-------------------|----------------|-----------------|-------------|
| Deploy frequency | 1.8/week | 3/week per service | 1+/day per service | CI/CD pipeline metrics |
| Lead time (commit → prod) | 12.3 days | 5 days | < 2 days | GitHub PR merge → deployment timestamp |
| Change failure rate | 18.2% | 12% | < 8% | Hotfix deploys / total deploys |
| MTTR | 4.2 hours | 1.5 hours | < 30 min | PagerDuty incident duration |
| P99 API latency | 820ms | 400ms | < 200ms | Prometheus histograms |
| Build time (CI) | 47 minutes | 15 min (per service) | < 8 min | GitHub Actions duration |
| Test flakiness | 6.8% | 3% | < 1% | CI failure rate excluding code bugs |
| Developer satisfaction (survey) | 5.8/10 | 7.0/10 | 8.0/10 | Quarterly engineering survey |

---

## Team Structure — Current vs Target

### Current (Monolith)
- **4 feature squads** (3-6 engineers each) — all deploy the same artifact
- **1 SRE team** (3 engineers) — owns all infrastructure
- Coordination overhead: O(n²) — every team's changes can break every other team's code
- Deploy queue: teams wait for "deploy slots" — informal but real bottleneck

### Target (Service-Oriented)
| Squad | Services Owned | Size | On-Call |
|-------|---------------|------|---------|
| Identity | User Service, Auth | 4 eng | Yes — own pager |
| Commerce | Order Service, Payment Service | 6 eng | Yes — own pager |
| Discovery | Search Service | 3 eng | Shared with Platform |
| Platform | Notification Service, API Gateway, shared libs | 4 eng | Yes — own pager |
| Data | Analytics Service | 3 eng | Shared with Platform |
| SRE | Kafka, Kubernetes, monitoring, incident response | 3 eng | Yes — primary escalation |

**Key principles:** You build it, you run it. Each squad owns their service's SLOs, deploys independently, and carries their own pager. SRE provides platform, not babysitting.

---

## Architecture Decision Record — Why Strangler Fig, Not Rewrite

This section is the narrative rationale. One slide should present this as dense text — this is what the VP Eng will reference when justifying the approach to the CTO.

**Why not a full rewrite?** We evaluated three approaches in a 2-week spike (April 14-25, 2026). The rewrite option (greenfield in Go + gRPC) was estimated at 14-18 months with a team of 8 — during which the existing monolith still needs maintenance, meaning we'd effectively run two systems. Netscape, Basecamp, and dozens of case studies show rewrites take 2-3x longer than estimated and frequently fail to ship. Our monolith has 340K lines of tested, working business logic. Rewriting it means re-discovering every edge case the hard way.

**Why not just modularize the monolith?** We tried this in Q4 2025. The "Atlas Modular" initiative spent 6 weeks extracting Django apps into separate packages with defined interfaces. Result: merge conflicts dropped 30% but deploy frequency didn't improve at all because we still ship one artifact. The 47-minute CI build didn't get faster. The blast radius didn't shrink. Modularization addresses code organization but not the deployment coupling that's our actual bottleneck.

**Why Strangler Fig specifically?** It's the only pattern that lets us ship incremental value while de-risking. Each extracted service runs in production alongside the monolith — we can validate it works before cutting over traffic. If a service extraction goes wrong, we route traffic back to the monolith. The gateway gives us a clean routing layer. The risk is bounded to one service at a time, not the entire system. We've already proven the pattern works: the User Service extraction in the Q1 2026 prototype took 3 weeks, ran in parallel for 2 weeks, and cut over with zero customer-facing incidents. That gives us confidence the approach scales to the harder extractions (Orders, Payments).

**What we're NOT doing:** We are not extracting every module. The Reporting and Admin modules stay in the monolith indefinitely — they're low-traffic, low-change, and the cost of extraction exceeds the benefit. The goal is not "zero monolith" — it's removing the modules that cause deployment friction for the teams that ship the most.

**Timeline confidence:** 65% confident in the 6-month timeline for Phase 1-2 (core extractions). 40% confident in Phase 3 completing on time — Analytics migration depends on ClickHouse adoption, which has an unknown learning curve. We've padded Phase 3 by 2 weeks but may need to extend.

---

## Slide Suggestions

6 slides, dense. Mix tables AND text — not everything is a chart:

1. **Title + Current State** — system stats (340K LOC, 22 engineers, pain metrics table)
2. **Target Architecture** — service boundary table + architecture diagram
3. **Why Strangler Fig** — the architecture decision narrative above as dense text with key decisions bolded. This tests text-wall handling.
4. **Migration Plan** — 6-month phased timeline with tasks, owners, risks per phase
5. **Risk Matrix + Monitoring** — severity-coded risk table + observability stack
6. **Success Metrics + Team** — before/after targets table + squad ownership table
