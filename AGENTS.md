# AGENTS.md

## SlideGet fork: mandatory engineering standard

These requirements apply to every change maintained for SlideGet, including
core, preview, editor, tests and tooling. Read the upstream guide below as well.
For SlideGet integration, the canonical JSON Deck/DeckOp model is the source of
truth; mutable OOXML and SVG are projections. Preserve upstream OOXML fidelity
and public compatibility. Do not import SlideGet server internals into this fork.

### Types and runtime contracts

- New or changed code must satisfy strict TypeScript, including unchecked index
  access, exact optional properties, exhaustive unions, explicit override,
  handled unknown errors, no implicit returns and no switch fallthrough. Keep
  public declarations accurate and test important type contracts with negative
  compile-time cases. Svelte components use TypeScript and typed props/events.
- Do not add `any`, unsafe casts/double casts, non-null assertions, `@ts-ignore`,
  `@ts-nocheck`, disabled strict options or broad lint suppressions. Narrow with
  runtime checks and discriminated unions. A necessary third-party typing
  exception belongs in one documented adapter with a regression test; repair
  our own inaccurate types rather than casting around them.
- External data starts as `unknown`: validate worker/postMessage traffic,
  persisted state, clipboard/HTML, imported ZIP/XML, public mutation payloads
  and asynchronous responses before use. Validate nested values, finite numbers,
  sizes/counts, discriminants, version, identities and relationships. Mutation
  DTOs whitelist writable fields and reject unknown fields. Use schemas where
  available; do not introduce a dependency just to repeat an already validated
  internal contract. Validation failure must leave state/history unchanged.
- Keep the public API small and typed. No private imports, hidden mutable graph
  escape hatches, unchecked generic response casts or duplicated source of truth.
  State ownership, buffer copying/aliasing and cache invalidation are explicit.

### Behavior, security and lifecycle

- Separate pure model/geometry/transaction logic from UI and transport. One
  user intent creates one atomic history operation. Test failure before prepare,
  during apply and during refresh; preserve undo/redo and the accepted model.
- Async operations need explicit cancellation, timeout, ordering and cleanup.
  Abort stale work, ignore obsolete replies, settle every pending promise and
  release workers, observers, listeners and object URLs. Cover destruction,
  reopen, concurrent requests and browser back/forward-cache restoration.
- Treat uploaded files, XML, SVG, links and clipboard HTML as untrusted. Enforce
  resource limits, relationship targets, supported protocols and sanitization at
  the owning boundary; reject unsupported capabilities rather than silently
  dropping data. No raw user content, prompts, credentials or provider payloads
  in logs or performance telemetry.
- Accessibility is behavior: visible focus, keyboard parity, correct semantics,
  focus/caret restoration, IME/composition and reduced motion need verification.
  A visually hidden control must not remain interactively focusable.

### Tests and proof

- Every changed behavior needs the smallest meaningful automated regression
  test. Reproduce a bug with a failing test before fixing it. Cover valid and
  invalid input, boundaries, unsupported features, stale versions, cancellation,
  rollback, duplicate delivery and concurrency where applicable. Do not write
  tests that merely repeat implementation statements.
- Use unit/property tests for pure algorithms and mutation journals; integration
  tests with real parser/serializer/package parts for projection and save/reload;
  real-browser component/flow tests for pointer/keyboard, selection, drag cancel,
  rich text/IME, groups, copy/paste, shared undo/redo, AI Accept and save/reopen.
  Mocked DOM, clocks or storage do not prove real-browser or file interoperability.
- Changed OOXML output needs byte/semantic round-trip assertions, relationship
  and identity checks, deterministic-output checks and the applicable schema/SDK
  validator. Changed painting/cache paths need original-output fidelity checks,
  dependency invalidation tests and browser DOM/layout verification. Test all
  fast paths against the uncached/reference path, including theme/font changes.
- Report coverage for critical changed model/history/validation modules; aim for
  at least 90% lines/functions without using the number to replace scenario
  coverage. Tests and tooling meet the same typing/quality rules as runtime code.
- Before completion run format, lint, types, applicable unit/integration/browser
  suites and build, plus relevant performance/bundle/fidelity gates. Review the
  actual outputs. Do not call unavailable, skipped or unexecuted checks passing.

### Performance is a correctness requirement

- Prefer lazy work: mount/render visible slide previews with bounded overscan;
  defer hidden panels and optional tools; start workers and exports on demand;
  stop idle measurement/animation loops. Do not defer work if it breaks command
  atomicity, accessibility, save completeness or visible correctness.
- Normal edits/undo/redo must process the affected object/slide/parts, preserve
  unchanged handles/DOM and avoid whole-deck XML scans, ZIP serialization,
  cloning and parsing. Batch serialization at the mutation boundary. Caches must
  have explicit dependency revisions, memory/count bounds and lifetime cleanup.
  Offload genuinely heavy serialization where useful, including cancellation.
- Measure p95 for every meaningful operation: live frame work, commit, store,
  undo, redo, text/group/clipboard, AI Accept, open and save. Keep raw samples,
  sample counts, metric boundaries, environment and reviewed source pins. Do
  not replace p95 with a single duration, average or an unlabelled proxy.
- The SlideGet 60 slides × 40 rich-text objects gate is: main-thread frame work
  p95 <=16ms; bridge commit p95 <=10ms; store p95 <=2ms; native command-to-next-rAF
  commit/undo/redo p95 <=50ms; no editing task >50ms; no whole-deck history
  rebuild. Display cadence and compositor presentation are different metrics.
  Preserve the same workload, raw slow samples and existing bundle-size gates.
- Compare against a versioned accepted baseline: >20% p95 regression fails.
  Missing required samples/infrastructure/baselines fail, not skip. Measure cold
  open separately, retain long-task and heap evidence, and investigate allocation,
  GC, DOM/layout and algorithm costs. Low-level optimization is welcome when a
  profile proves the bottleneck and tests prove behavior/fidelity unchanged.
  Never raise budgets, reduce workload or filter slow samples to pass.

### Maintainability and enforcement

- Prefer simple explicit algorithms, small modules and stable ownership over
  clever abstraction. New dependencies require a one-line justification. New
  custom scripts/configuration are TypeScript (JSON where required); portable
  commands use the pinned package manager without OS-specific shell syntax.
  Preserve existing upstream source conventions and Svelte; no blanket rewrite.
- Keep format/lint/type/test/build gates enabled. Type-aware unsafe-value,
  promise, exhaustive-union, import-boundary and framework-reactivity checks
  belong in the owning quality configuration. New tests must be discovered in
  CI; missing expected suites/infrastructure must fail. No success placeholders,
  `continue-on-error`, test exclusions or weakened thresholds to hide failures.
- These rules are requirements, not a claim that inherited upstream code already
  satisfies every rule. Record an exact-head baseline for inherited failures and
  configuration gaps; fix touched unsafe paths, do not add new debt or silently
  exempt new code. A wider strictness/tooling migration is an explicit tracked
  task, with an owner and a fail-on-new-regression gate, not an undocumented skip.
- Before reporting completion, inspect the diff and obtain independent review
  for public APIs, model/history/cache changes. Include changed files, dependency
  reasons, exact commands/results, performance before/after, known failures and
  remaining limits. Add a changeset for user-visible changes, preserve NOTICE
  and keep fork patches focused enough to rebase on upstream.

The upstream guide remains mandatory below. This fork section is authoritative
for SlideGet-specific quality and canonical-model decisions.

The upstream instructions live in [CLAUDE.md](CLAUDE.md). **Read it before
changing anything here** — it is not Claude-specific. It carries the project
overview, the OOXML rules that decide what correct output is, the public-API
constraints, and the review standard every change is held to.

Keep shared upstream guidance in `CLAUDE.md`; do not duplicate it here. All agents,
including those loading `CLAUDE.md` automatically, must read this fork's quality
section above before making SlideGet changes.

A few things worth knowing before you start:

- **Do not author slide decks with your own presentation tooling.** This project
  is that tooling. Decks are written as TSX against `@office-kit/pptx-dsl` and
  exported with `@office-kit/pptx-dev`; see the "TSX presentation authoring"
  section of CLAUDE.md. A generated deck project carries its own guide under
  both names.
- **The package manager is pnpm**, pinned by `packageManager` in
  `package.json`. Use `pnpm`, not `npm` or `yarn`.
- **Quality gates**, in the order CI runs them:

  ```sh
  pnpm format:check && pnpm lint && pnpm typecheck && pnpm test && pnpm build
  ```

  The DSL package type-checks against the built output of the core and preview
  packages, so `pnpm --filter @office-kit/pptx-dsl typecheck` is only meaningful
  after those have been built. Running it against a stale `dist/` reports fewer
  errors than CI will.

- **User-visible changes need a changeset** (`pnpm changeset`). See the
  "CHANGELOG is for users" section of CLAUDE.md for what belongs in one.
