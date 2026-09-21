---
name: own-website-audit
description: Comprehensive audit for landing pages and websites across advertising and legal compliance, SEO, marketing conversion quality, design/UX, accessibility, technical health, and analytics/measurement. Use when asked to review, audit, evaluate, QA, or preflight a website, LP, service site, corporate site, article site, recruiting site, or form-based conversion flow before or after launch. 「LPを評価して」「このコピーは刺さるか」「LPを公開前に見て」「リデザインの良し悪しを判定して」でも使う（LP 評価の専用 skill を 2026-09-21 に本 skill へ統合した）。審美とレイアウトは `scripts/measure_lp.mjs` で**決定論的に実測**し、**それが緑になるまで独立レビュアーへ進まない**。使わない場面: デザインの新規生成（web は own-web-design、スライドは own-pptx-build）、検査の振り分け（own-design-route）、合格まで直して再評価するループ（own-output-eval）。**決定論の実測だけを回したい場合**は本 skill を経由せず `scripts/measure_lp.mjs` を直接叩く。
---

# Website Audit

Use this skill to audit websites and landing pages comprehensively. The goal is to find launch-blocking and performance-limiting issues across compliance, marketing, SEO, UX, accessibility, technical quality, and measurement.

## Inputs

Accept any of these targets:

- Public URL, staging URL, or local dev server URL
- Static HTML or built site files
- Source files for a page, app route, component, or template
- Screenshots when browser access is unavailable

When provided, also read only the relevant context files:

- PRD, campaign brief, ad copy, keyword list, target audience, or conversion goal
- Brand/design guidelines
- Legal, privacy, pricing, offer, or industry-specific requirements
- Ad platform requirements such as Google Ads policy notes
- Existing analytics/GTM/ads conversion specifications

If context is missing, continue with generic website best practices and mark assumptions clearly.

## Core Workflow

1. Identify the site type, primary audience, main conversion goal, traffic sources, and target pages.
2. Inspect the rendered experience when a URL is available. Test mobile, tablet, desktop, and wide desktop where practical.
3. Read source files only as needed to validate findings or propose concrete fixes.
4. **決定論を先に実測する（LP・単一ページのとき必須）。**

   ```bash
   node <skill>/scripts/measure_lp.mjs <url> --viewports 375,1280 --schemes dark,light --json /tmp/lp.json
   ```

   合格閾値は `references/aesthetic-thresholds.md`。溢れ・重なり・切れ・コントラスト・
   不安定な余白は、人が見る前に機械で落とす。

5. **ゲート — 決定論が緑になるまで、非決定論へ進まない。**

   **失格が 1 件でも残っているうちは、独立レビュアーを走らせない。**
   理由は費用と信頼性の両方にある。レビュアーは 1 人あたり数万トークンを使う一方、
   表示が壊れたページを採点しても出てくるのは表示不良の指摘だけで、
   **本来見たかったコピーと訴求の評価は汚染される**。

6. **マーケティング定性は独立レビュアーの項目別スコアで採点する。**
   `references/marketing-rubric.md` の 10 項目を、**そのページを作っていない**
   レビュアーに採点させる。作った本人の自己採点は甘くなる。

7. Evaluate every required category below.
8. Report findings by severity with evidence, impact, recommended fix, and verification method.
9. Separate confirmed issues from assumptions or items needing human/legal review.

## Required Audit Categories

### Advertising and Legal Compliance

Check for risks that could cause ad disapproval, user deception, or legal exposure:

- Platform policy fit for Google Ads and other declared ad platforms
- Prohibited or restricted category claims based on the site topic
- Exaggerated, absolute, fear-based, or unverifiable claims
- Unsupported performance numbers, rankings, testimonials, guarantees, or before/after claims
- Misleading redirects, disguised ads, missing destination relevance, or mismatch with ad copy
- Clear operator/business identity, contact path, privacy policy, terms, pricing, conditions, disclaimers, and eligibility limits
- Proper handling of sensitive categories such as health, finance, employment, housing, legal, adult, politics, and personal hardship

Do not present legal conclusions as legal advice. Flag legal-sensitive items as review risks when a qualified review is needed.

### Marketing and Conversion Quality

Check whether the page can convert the intended traffic:

- Match between traffic intent/ad copy/keywords and above-the-fold message
- Clear offer, audience fit, differentiation, and trust proof
- CTA visibility, specificity, repetition, and continuity through the page
- Form length, field friction, validation clarity, confirmation state, and fallback contact path
- Objection handling, social proof, process explanation, pricing clarity, and next-step expectations
- Mobile conversion path and tap target ergonomics

### SEO

Check discoverability and search intent alignment:

- Title, meta description, canonical, robots, sitemap, indexability, and URL hygiene
- H1/H2 structure, semantic sectioning, internal links, breadcrumbs where relevant
- Search intent match, topical coverage, originality, and content depth
- E-E-A-T signals: author/operator clarity, sources, dates, credentials, reviews, and trust pages
- Structured data opportunities and implementation risks
- Duplicate/thin content, hidden text, doorway-page patterns, and keyword stuffing

### Design and UX

Check visual quality and usability:

- Responsive behavior at mobile, tablet, desktop, and wide desktop
- Layout overflow, overlap, clipping, awkward wrapping, and unstable spacing
- Visual hierarchy, scanability, CTA prominence, navigation clarity, and wayfinding
- Typography, color, spacing, component consistency, and brand fit
- Form experience, error recovery, loading/empty states, and confirmation flows
- Avoid generic feedback; tie design findings to user comprehension, trust, or conversion impact

When available, use `web-design-reviewer` for visual/responsive inspection. If it is not available, use `web-design-guidelines`, `playwright-cli`, browser automation, screenshots, or source review as practical.

### Accessibility

Check the highest-impact WCAG-style issues:

- Accessible names for links, buttons, controls, icons, and form fields
- Keyboard access, visible focus, focus order, and dialog/flyout behavior
- Semantic HTML, heading order, landmark structure, lists, and tables
- Color contrast, non-color indicators, reduced motion, and readable text sizing
- Image alt text, media captions where relevant, and form error association

Use `fixing-accessibility` when detailed remediation guidance is needed.

### Technical and Measurement Quality

Check whether the site can be trusted, measured, and operated:

- Console errors, broken requests, broken links, redirects, 404s, mixed content, and SSL issues
- Core Web Vitals and performance risks: LCP, CLS, INP, image sizing, render-blocking resources, third-party scripts
- OGP/Twitter cards, favicon, manifest, localization tags where relevant
- GA4/GTM/ad conversion events, event naming, consent behavior, duplicate firing, and thank-you/lead state tracking
- Error monitoring, form submission reliability, spam protection, and data privacy boundaries

For React or Next.js code, use `vercel-react-best-practices` when performance, rendering, bundle size, or data-fetching risks are in scope.

## Severity Rubric

- `P0` Launch/ad-blocking: likely ad disapproval, serious legal/privacy risk, broken conversion path, severe accessibility blocker, or site unavailable.
- `P1` High impact: materially hurts conversion, SEO indexability, trust, measurement accuracy, or core mobile UX.
- `P2` Medium impact: weakens clarity, performance, accessibility, maintainability, or marketing effectiveness but does not block launch.
- `P3` Low impact: polish, minor consistency, optional SEO/UX enhancement, or future optimization.

## Output Format

Use this structure:

```markdown
# Website Audit Report

## Summary
- Target: <URL/files/pages>
- Site type: <LP/corporate/service/article/recruiting/form flow/unknown>
- Primary goal: <goal or assumption>
- Overall status: <Ready / Ready with fixes / Not ready>
- Top risks: <3-5 concise bullets>

## Findings
### [P0|P1|P2|P3] <short title>
- Category: <Compliance / Marketing / SEO / Design UX / Accessibility / Technical Measurement>
- Evidence: <specific page, element, text, screenshot, source file, or observed behavior>
- Impact: <why this matters>
- Recommendation: <specific fix>
- Verification: <how to confirm the fix>

## Positive Notes
- <Only include concrete strengths that affect trust, conversion, accessibility, SEO, or reliability.>

## Assumptions and Review Needed
- <State missing context, legal-review items, policy uncertainty, or data that could not be verified.>

## Suggested Next Checks
- <Browser/device tests, Lighthouse, Search Console, Ads preview, GTM debug, legal review, etc.>
```

Keep the report evidence-based. Do not invent policy failures or metrics that were not observed. If browsing, live testing, or source access is unavailable, state the limitation and audit the available artifact.

## External Skill Notes

- Prefer installed, trusted local skills over low-install external skills.
- `web-design-reviewer` is a recommended optional helper for design and responsive inspection if installed.
- Low-install advertising or LP audit skills may be used only as reference material after reading their instructions; do not rely on them as authoritative policy sources.
- For up-to-date ad platform policy, official platform documentation should be checked when the decision affects launch, spend, or account risk.
