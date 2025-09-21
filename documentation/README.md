# Documentation - Document Processing API

## 📚 Two Main Documents (Everything Consolidated)

### 1. `CODE_MAP_AND_MCP_REFERENCE.md` (10KB)
**Purpose**: Technical reference for code fixes and resuming work
- Exact code locations for all fixes (file + line numbers)
- MCP context for project resume
- Testing commands and deployment instructions
- Requirements & dependencies
- API endpoints reference
- Performance statistics
- VML structure details

### 2. `STRATEGIC_ROADMAP_AND_FIXES.md` (11KB)
**Purpose**: Strategic planning and issue tracking
- Current state analysis (Word COM vs ZIP)
- Phase 1-4 roadmap
- All critical fixes needed
- Deployment cost analysis ($50 vs $500/month)
- Frontend improvements
- Installation & setup commands
- Thread findings summary

---

## 📁 Chat History Archive

Detailed conversation threads kept for reference:
- `chat_history/thread 1/` - Initial setup and discovery
- `chat_history/thread 2/` - VML textbox limitation discovery
- `chat_history/thread 3/` - Final implementation and solutions

---

## 🎯 When Resuming Project

**Start with**: `CODE_MAP_AND_MCP_REFERENCE.md`
- Check "Resume Instructions" section
- Review "Current Crossroads" diagram
- Follow code map to specific fixes

**For planning**: `STRATEGIC_ROADMAP_AND_FIXES.md`
- See Phase 1-4 breakdown
- Check deployment options
- Review cost analysis

---

## Key Decision: Fix ZIP Approach First!

**Why**: Saves $450/month on deployment ($50 vs $500)

**Main Issue**: ZIP approach has file corruption, Word COM needs expensive Windows VM

**Strategic Path**: Fix ZIP → Deploy to Cloud Run → $50/month