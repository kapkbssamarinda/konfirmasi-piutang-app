# Graph Report - konfirmasi-piutang-app  (2026-08-28)

## Corpus Check
- Corpus is ~13,409 words - fits in a single context window. You may not need a graph.

## Summary
- 64 nodes · 63 edges · 25 communities (6 shown, 19 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Vite & Project Configuration
- React UI & State Management
- Document Generation & Docxtemplater
- Spreadsheet Data Import & Parsing
- Form Validation & Detail Audit
- UI Styling & Auto-Animate
- Community 6
- Community 7
- Community 8
- Community 9
- Community 10
- Community 11
- Community 12
- Community 13
- Community 14
- Community 15
- Community 16
- Community 17
- Community 18
- Community 19
- Community 20
- Community 21
- Community 22

## God Nodes (most connected - your core abstractions)
1. `scripts` - 5 edges
2. `App()` - 4 edges
3. `@formkit/auto-animate` - 2 edges
4. `docxtemplater` - 2 edges
5. `file-saver` - 2 edges
6. `framer-motion` - 2 edges
7. `jszip` - 2 edges
8. `lucide-react` - 2 edges
9. `motion` - 2 edges
10. `pizzip` - 2 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Import Cycles
- None detected.

## Communities (25 total, 19 thin omitted)

### Community 0 - "Vite & Project Configuration"
Cohesion: 0.20
Nodes (9): name, private, scripts, build, dev, lint, preview, type (+1 more)

### Community 1 - "React UI & State Management"
Cohesion: 0.53
Nodes (4): App(), formatNominalTyping(), formatRupiah(), Icons

### Community 2 - "Document Generation & Docxtemplater"
Cohesion: 0.40
Nodes (5): @formkit/auto-animate, dependencies, @formkit/auto-animate, react-confetti, react-confetti

### Community 3 - "Spreadsheet Data Import & Parsing"
Cohesion: 0.67
Nodes (3): eslint, devDependencies, eslint

## Knowledge Gaps
- **31 isolated node(s):** `name`, `private`, `version`, `type`, `dev` (+26 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **19 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `dependencies` connect `Document Generation & Docxtemplater` to `Vite & Project Configuration`, `Form Validation & Detail Audit`, `Community 8`, `Community 9`, `Community 11`, `Community 12`, `Community 13`, `Community 14`, `Community 15`, `Community 16`, `Community 17`, `Community 18`?**
  _High betweenness centrality (0.546) - this node is a cross-community bridge._
- **Why does `devDependencies` connect `Spreadsheet Data Import & Parsing` to `Vite & Project Configuration`, `UI Styling & Auto-Animate`, `Community 6`, `Community 7`, `Community 10`, `Community 19`, `Community 20`, `Community 21`, `Community 22`?**
  _High betweenness centrality (0.415) - this node is a cross-community bridge._
- **What connects `name`, `private`, `version` to the rest of the system?**
  _31 weakly-connected nodes found - possible documentation gaps or missing edges._