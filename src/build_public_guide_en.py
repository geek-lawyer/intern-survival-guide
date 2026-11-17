"""Build the complete English public edition with the validated handbook design."""

from datetime import datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent.parent
BASE_BUILDER = PROJECT_DIR / "src" / "build_public_guide.py"
OUT = PROJECT_DIR / "dist" / "The Intern Survival Guide - Public Edition.docx"


# Reuse the exact page, style, list, table, bookmark and field helpers from the
# Chinese public builder.  Only the content block is replaced.
preamble = BASE_BUILDER.read_text(encoding="utf-8").split("# Cover", 1)[0]
preamble = preamble.replace(
    'OUT = PROJECT_DIR / "dist" / "实习生生存指南（公开阅读版）.docx"',
    'OUT = PROJECT_DIR / "dist" / "The Intern Survival Guide - Public Edition.docx"',
)
preamble = preamble.replace('paragraph.add_run("第 ")', 'paragraph.add_run("Page ")')
preamble = preamble.replace('paragraph.add_run(" 页")', 'paragraph.add_run("")')
preamble = preamble.replace(
    'header.add_run("实习生生存指南  ·  公开阅读版")',
    'header.add_run("GEORGE HU  ·  THE INTERN SURVIVAL GUIDE")',
)
preamble = preamble.replace('def quote(text, label="请记住"):', 'def quote(text, label="Remember"):')
preamble = preamble.replace('"返回目录"', '"Back to contents"')
namespace = {"__file__": str(BASE_BUILDER), "__name__": "__english_public_builder__"}
exec(compile(preamble, str(BASE_BUILDER), "exec"), namespace)
globals().update(namespace)


def chapter(number, title, lead, paragraphs=(), bullet_items=(), numbered_items=(), quote_text=None):
    doc.add_heading(f"Chapter {number}  {title}", level=1)
    p(lead, style="Lead")
    for text in paragraphs:
        p(text)
    if numbered_items:
        numbers(list(numbered_items))
    if bullet_items:
        bullets(list(bullet_items))
    if quote_text:
        quote(quote_text)


def add_contents_link(anchor, title, scope):
    para = doc.add_paragraph(style="List Bullet")
    add_internal_hyperlink(para, title, anchor, bold=True)
    r = para.add_run(" | " + scope)
    font(r, 9.5, color=MUTED)


# Cover
for _ in range(4):
    doc.add_paragraph()
k = doc.add_paragraph()
k.alignment = WD_ALIGN_PARAGRAPH.CENTER
font(k.add_run("INTERN SURVIVAL GUIDE"), 10, True, GOLD)
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_before = Pt(12)
title.paragraph_format.space_after = Pt(8)
font(title.add_run("The Intern\nSurvival Guide"), 30, True, NAVY)
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
font(sub.add_run("From student to reliable professional collaborator"), 15, color=TEAL)
byline = doc.add_paragraph()
byline.alignment = WD_ALIGN_PARAGRAPH.CENTER
byline.paragraph_format.space_before = Pt(12)
byline.paragraph_format.space_after = Pt(2)
font(byline.add_run("Author and editor | George Hu"), 10.5, True, NAVY)
contributors = doc.add_paragraph()
contributors.alignment = WD_ALIGN_PARAGRAPH.CENTER
contributors.paragraph_format.space_after = Pt(22)
font(contributors.add_run("Built with interns, mentors and repeated real-world practice"), 9.5, color=MUTED)
quote("This guide does not help you wait out an internship. It answers three questions: How do I choose a direction? How do I build ability? How do I earn trust?", "Purpose")
doc.add_paragraph().paragraph_format.space_after = Pt(60)
meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
font(meta.add_run("PUBLIC READING EDITION  ·  JULY 2026"), 9.5, color=MUTED)

doc.add_page_break()
doc.add_heading("Copyright and version", level=1)
p("When authorship, version and permission are unclear, every later copy makes responsibility less clear.", style="Lead")
add_table(["Item", "Information"], [
    ("Title", "The Intern Survival Guide - Public Edition"),
    ("Author and editor", "George Hu (Hu Zexian)"),
    ("Practice base", "Lessons from multiple cohorts of interns and mentors"),
    ("First edition", "July 2026"),
    ("Current edition", "July 2026 public reading edition; use the latest author-published version"),
], [1.7, 4.85])
p("Copyright © 2026 George Hu. Text licensed under CC BY-NC-SA 4.0.", italic=True, after=4)
p("The original structure, writing and compilation are protected by copyright. Linked third-party articles remain the property of their respective authors and publishers.")
p("The text of this edition is released under the Creative Commons Attribution-NonCommercial-ShareAlike 4.0 licence (CC BY-NC-SA 4.0): you may copy, share, translate and adapt it with credit to the title, author, edition and source, provided adaptations carry the same licence. Commercial use, including commercial publishing and paid training, needs the author's written permission.")
p("This is a training and management guide, not legal advice. Every team should adapt its tools and procedures to its own rules, technology, risk profile and applicable law.")

doc.add_page_break()
doc.add_heading("Contact, updates and corrections", level=1)
p("A handbook starts becoming obsolete when it stops accepting corrections.", style="Lead")
add_table(["Channel", "Information"], [
    ("Email", "hu.zexian@qq.com"),
    ("WeChat public account", "没事多吃核桃"),
    ("Website", "brbrhu.com"),
], [1.7, 4.85])
qr_para = doc.add_paragraph()
qr_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
qr_shape = qr_para.add_run().add_picture(str(OFFICIAL_ACCOUNT_QR), width=Inches(2.8))
qr_shape._inline.docPr.set("title", "WeChat public account QR code")
qr_shape._inline.docPr.set("descr", "QR code for the WeChat public account 没事多吃核桃")
caption = doc.add_paragraph()
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
font(caption.add_run("Scan to follow the public account and send corrections"), 9.5, color=MUTED)

doc.add_page_break()
doc.add_heading("How to use this guide", level=1)
p("A handbook that only controls interns, while placing no duty on managers, is merely a command list.", style="Lead")
p("Read it once on your first day. During the first week, use it alongside confidentiality, computer, shared-storage and equipment training. When a real question appears, go first to Part III or the appendices. Review one question each week.")
p("The legal examples reflect the guide's origin. The operating principles apply more broadly: define the purpose, protect sensitive information, work from evidence, test a small sample, record the result and leave behind a reusable method.")

doc.add_heading("What kind of team is this for?", level=2)
add_table(["Team condition", "Learning opportunity", "Management duty"], [
    ("Real work with real consequences", "Learn how facts, people, deadlines and professional judgment interact.", "Open work by risk level; explain where the task fits and review the result."),
    ("Mentors with different backgrounds", "Observe more than one way to solve a problem.", "Turn expertise into tasks, examples and consistent feedback."),
    ("Computers, shared storage and AI", "Build file, retrieval, automation and information-security habits.", "Teach access and workflow from zero; equipment does not teach itself."),
    ("Optional media or production tools", "Translate complex work into clear public communication.", "Protect accuracy, privacy, copyright and release approval."),
    ("Previous cohorts and templates", "Start from accumulated lessons rather than repeating every mistake.", "Keep honest feedback, remove obsolete rules and maintain the knowledge base."),
], [1.35, 2.55, 2.65])

doc.add_heading("The promise must run both ways", level=2)
bullets([
    "Interns take responsibility for honesty, confidentiality, careful work and early risk reporting.",
    "Managers state the purpose, output, standard, deadline, owner and review method.",
    "Once basic work is stable, interns receive progressively harder and more meaningful tasks.",
    "Strictness targets observable behavior and work quality, never humiliation or vague rejection.",
    "Both sides may decide whether the internship remains a good fit; gratitude does not erase reasonable boundaries.",
])

doc.add_page_break()
toc = doc.add_heading("Find the problem before the answer", level=1)
add_bookmark(toc, "toc")
p("In the electronic edition, click a part to jump to it.", italic=True, after=5)
add_contents_link("part1", "Part I  Purpose", "Chapters 1-7")
add_contents_link("part2", "Part II  Mutual growth", "Chapters 8-12 and the mentor chapter")
add_contents_link("part3", "Part III  Getting operational", "Chapters 13-16")
add_contents_link("part4", "Part IV  Professional delivery", "Chapters 17-23")
add_contents_link("part5", "Part V  Direction", "Chapters 24-26")
add_contents_link("appendix", "Appendices", "Eleven reusable tools")

doc.add_heading("Start from your current question", level=2)
add_table(["If you are asking...", "Read first"], [
    ("How do I use the work computer or shared storage?", "Part III, then Appendices 1 and 2"),
    ("How do I accept a task without guessing?", "Chapter 13 and Appendix 3"),
    ("How do I send or submit something safely?", "Part IV and Appendix 7"),
    ("How may AI be used?", "Chapter 23 before entering any material"),
    ("How should I manage an intern?", "Part II and Appendices 8 and 11"),
    ("How do I review the internship and choose a direction?", "Part V and Appendices 4, 9 and 10"),
], [3.7, 2.8])

section_break("Part I  Purpose", "Why this internship deserves serious attention", "part1")

chapter(1, "From standard answers to real consequences",
    "School can grade an answer. Real work changes what happens to people.",
    paragraphs=(
        "A task may arrive as incomplete facts, an urgent deadline and a person who cannot afford casual mistakes. An intern may not make the final decision, but remains responsible for every fact copied, every page moved and every uncertainty concealed.",
        "Professionalism is not mystery. It is the habit of staying honest, careful, restrained and reliable when another person's important matter is in your hands.",
    ), quote_text="The internship starts before the first word is written: first ask who that word may affect.")

chapter(2, "Ask who will use the deliverable",
    "Completing an instruction is the minimum; solving the user's problem is the purpose.",
    paragraphs=("A file is organized so the next person can understand it. Research is done so someone can make a better decision. A scan is useful only if it is complete, readable and placed where the next step can find it.",),
    numbered_items=(
        "What final outcome is this task meant to support?",
        "Who will use my work?",
        "What will that person need first?",
        "Which mistakes could cause serious consequences?",
    ), quote_text="Before delivery, explain the purpose, user, standard and risks in one minute.")

chapter(3, "Do not design your life around the minimum",
    "What looks fair in the short term may sell your most valuable learning time too cheaply.",
    paragraphs=(
        "Compensation is easy to count, but it is not the only return. Count real tasks, specific feedback, independent delivery and professional judgment as well. A low-paid internship with only consumption is expensive; one that steadily builds capability may still be valuable.",
        "Long-term opportunity must be visible in today's task, the person who will review it and the responsibility that follows good work. 'Endure now, perhaps later' is not a development plan.",
        "Long-term thinking never excuses unreasonable arrangements. A team must also provide learning opportunities, feedback and boundaries.",
    ), quote_text="Do not obsess over a small temporary loss. Do judge whether the experience makes you more capable, discerning and trustworthy.")

chapter(4, "Small work decides who receives bigger work",
    "Important work is rarely given to someone who despises the small steps that protect it.",
    paragraphs=(
        "File names, page order, attachments, meeting setup and follow-up notes look small only when they work. When they fail, everyone downstream pays.",
        "Reliability is cumulative. Each correct small delivery lowers the cost of trusting you with the next one.",
    ), bullet_items=(
        "Finish the entire task, including naming, saving, checking and reporting.",
        "Correct obvious errors before asking another person to review.",
        "If a step feels repetitive, ask what risk the repetition is controlling.",
    ))

chapter(5, "Do not confuse activity with growth",
    "A full calendar proves movement, not progress.",
    paragraphs=("Growth appears when the second attempt is better: fewer corrections, clearer questions, faster retrieval, better risk recognition and a reusable method left behind.",),
    numbered_items=(
        "What can I now do without supervision?",
        "Which error will my new checklist prevent?",
        "What judgment has become more precise?",
        "What useful asset did I leave for the team?",
    ))

chapter(6, "An internship is mutual selection, not performance",
    "The team observes how you work; you observe what kind of work you can live with.",
    paragraphs=(
        "Do not perform enthusiasm while hiding confusion. Give honest information about your availability, goals and limits. Ask to see the real task, review style and working rhythm.",
        "The internship is evidence for a career decision. Notice whether you can repeatedly face facts, deadlines, uncertainty, revision and responsibility - and whether the team develops people rather than merely consumes them.",
    ))

chapter(7, "Do not trust memory: build checklists",
    "A checklist respects human limits; it does not insult intelligence.",
    paragraphs=(
        "A useful SOP has four layers: explanation for the novice, a short checklist for the familiar, a release gate for non-negotiable risks and a log for later learning.",
        "Use a small checklist for low-risk work. Add operator, reviewer, red lines and an explicit hold decision when the task is irreversible, external, time-critical or confidential.",
    ))
add_table(["Layer", "Question", "Typical form"], [
    ("Guide", "How does a newcomer understand the work?", "Purpose, steps, examples, common mistakes"),
    ("Checklist", "How does a familiar person avoid omissions?", "Short observable checks"),
    ("Release gate", "What must stop the work?", "Red lines, reviewer, hold option"),
    ("Log", "What happened and what should change?", "Time, quantity, exception, response"),
], [1.2, 3.05, 2.35])
quote("First block failure; then pursue speed.", "Checklist principle")

section_break("Part II  Mutual growth", "What interns and managers owe each other", "part2")

chapter(8, "Initiative is not unauthorized decision-making",
    "Mature initiative means understanding first, researching next and asking with a proposed path.",
    paragraphs=("Escalate immediately when the issue concerns direction, deadlines, another person's interests, missing material, conflicting instructions, a mistaken send or possible disclosure.",))
add_table(["Level", "Action", "Examples"], [
    ("Self-check", "Read the handbook, shared knowledge base, prior records and reliable public material.", "File naming, basic format, routine equipment"),
    ("Ask nearby", "Ask another intern, mentor or direct task owner; explain what you already checked.", "Template location, local operation, partial ambiguity"),
    ("Escalate now", "Report to the accountable professional; do not guess or wait.", "Deadline, strategy, sensitive data, missing or conflicting material"),
], [1.1, 3.6, 2.0])

chapter(9, "Care, effort and commitment: ask three questions",
    "Values become useful only when they can question a real action.")
doc.add_heading("Care: Do I understand what I am doing?", level=2)
bullets([
    "Can I state the purpose, user, standard and main risk?",
    "Have I separated fact, inference and uncertainty?",
    "Did I read the source and sample before starting?",
])
doc.add_heading("Effort: Have I taken enough action for the result?", level=2)
bullets([
    "Did I make a small sample before scaling the work?",
    "Did I leave a record of obstacles and attempted solutions?",
    "If timing changed, did I report early rather than at the deadline?",
])
doc.add_heading("Commitment: Have I remembered the people behind the file?", level=2)
bullets([
    "Would I handle this the same way if it were my own important matter?",
    "Can the next colleague use my result without reconstructing my thinking?",
    "Did I protect the dignity, privacy and time of the people affected?",
])
add_table(["Moment", "Care", "Effort", "Commitment"], [
    ("Before", "Do I understand?", "What sample will I test?", "Who will be affected?"),
    ("During", "What changed?", "What obstacle needs action?", "Who needs an update?"),
    ("After", "What did I learn?", "What must be checked again?", "What can I leave for others?"),
], [1.0, 1.9, 1.9, 1.7])

chapter(10, "Do not fear revealing what you do not know",
    "A hidden gap is more dangerous than an admitted one.",
    paragraphs=("Ask a narrow question with context: what you were asked to do, what you checked, where the uncertainty remains and what answer you currently think is most likely.",),
    quote_text="A strong question reduces uncertainty; it does not transfer all thinking to someone else.")

chapter(11, "Translate criticism into the next action",
    "Feedback is useful when it changes the next observable step.",
    paragraphs=("Managers should name the missing step, its impact, the check to add and the resubmission time. Interns should restate the correction and apply it to similar work, not only the marked line.",),
    bullet_items=(
        "Separate a mistake from a judgment about the person.",
        "Ask whether the standard was clear before the work started.",
        "Record recurring errors as a checklist or training item.",
        "When the second version improves, reduce supervision rather than creating permanent dependence.",
    ))

doc.add_heading("Mentor chapter  Training an intern is also a management test", level=1)
p("Interns are cultivated; employees are selected. Cultivation reveals the manager's ability to define, teach, observe and improve.", style="Lead")
doc.add_heading("The person who assigns the task owns the teaching", level=2)
p("A task owner may ask others for help, but may not disappear after giving an incomplete instruction. One task needs one accountable owner and one final standard.")
doc.add_heading("Managers must manage themselves first", level=2)
bullets([
    "Follow the same confidentiality, timing, filing and checking rules you require from interns.",
    "Do not use status, fear or vague disapproval as a substitute for a clear standard.",
    "Do not turn care into a debt that demands loyalty, overtime, privacy or silence.",
    "Explain different treatment with task-relevant reasons that can be stated openly.",
])
quote("Strictness without care becomes obedience. Care without boundaries becomes debt. Neither is credible when the manager is exempt from the same rules.", "Manager's test")
doc.add_heading("Measure development, not gratitude", level=2)
add_table(["Evidence", "What improvement looks like"], [
    ("Task quality", "Fewer corrections; errors move from basic to genuinely difficult."),
    ("Independence", "The intern can find the entry point, make a sample and report risk."),
    ("Judgment", "Facts, assumptions, sources and uncertainty are separated."),
    ("Team asset", "A new checklist, template or knowledge entry remains usable."),
], [2.0, 4.5])
doc.add_heading("Use a calibration meeting when shared understanding breaks", level=2)
p("If several people say an intern is 'hard to manage,' convert the label into facts. Clarify purpose, roles, first-week standards, questions, available support and the next sample. Then review actual change three to five workdays later.")
doc.add_heading("Recruitment chooses an entry point; development shapes the path", level=2)
p("A resume estimates a starting point. A short work sample, learning speed, revision after feedback and reliable delivery provide better evidence. Do not use school, major or one early mistake as a permanent label.")

chapter(12, "Confidentiality is the first gate",
    "Skill can be developed gradually. Confidentiality begins on day one.",
    bullet_items=(
        "Do not photograph, copy, forward or remove sensitive material without permission.",
        "Store work only on approved computers, shared storage and work accounts.",
        "Separate work communication from personal social sharing.",
        "Lock the screen and control paper originals when leaving the desk.",
        "Do not enter client, case, identity or internal-strategy material into AI without explicit permission.",
        "Report a mistaken send, lost material or possible disclosure immediately; do not conceal it.",
    ), quote_text="Confidentiality does not depend on whether a matter looks important. It depends on the trust placed in us.")

section_break("Part III  Getting operational", "Master the environment before taking on professional work", "part3")

doc.add_heading("First-week training map", level=1)
p("The first week is not a production contest. Remove four sources of rework: lost files, wrong formats, unfamiliar equipment and unclear confidentiality boundaries.", style="Lead")
add_table(["Day", "Training", "Evidence of completion"], [
    ("1", "Confidentiality, work accounts, computer basics, file manager, lock screen, screenshots", "Create and locate the required folder; annotate a screenshot; lock the screen"),
    ("2", "Naming, copy/move, version saving, shared-storage access, paths, PDF basics", "Process and file a practice set in the approved shared directory"),
    ("3", "Scan, print, OCR, combine/split PDF, supplies and basic troubleshooting", "Produce a checked practice packet"),
    ("4", "Task restatement, small samples, updates, calendar, handbook and question order", "Complete a task card and a concise progress update"),
    ("5", "Integrated exercise and first-week review", "Complete one workflow from receipt to verified filing"),
], [0.7, 3.75, 2.15])

chapter(13, "Accept a task in five steps",
    "Do not confirm only the deadline; confirm what counts as complete.",
    numbered_items=(
        "Purpose: What problem must the work solve and who will use it?",
        "Output: Document, spreadsheet, PDF, physical packet or oral briefing?",
        "Standard: Template, format, name, location and reference sample?",
        "Timing: Final deadline and intermediate review point?",
        "Restatement: Explain the task in your own words and resolve differences now.",
    ))

chapter(14, "Make the environment ready to work",
    "An office does not enter a working state by itself.",
    paragraphs=("At the start, check the space, shared equipment, supplies and sensitive material. At the end, file work, reset equipment, clear waste and confirm physical security. A rota assigns actions; it does not cancel everyone's duty to notice and report problems.",))

chapter(15, "Control the work computer before chasing efficiency",
    "Not knowing a new system is normal. Turning the same gap into repeated rework is not.",
    paragraphs=(
        "Every team may use a different operating system or storage provider. Learn the local equivalents for application switching, locking, search, screenshots, file creation, rename, move, preview and safe device removal.",
        "The first file rule is universal: never overwrite the source. Save a working copy, convert it, open the result and confirm the actual path.",
    ))
add_table(["Task", "What must be demonstrated"], [
    ("File manager", "Create, search, sort, rename, copy, move and confirm the real path"),
    ("Shared storage", "Sign in or mount, find the approved directory, keep the source and verify upload"),
    ("PDF", "Create, combine, split, rotate, OCR, paginate and reopen the result"),
    ("Security", "Use approved accounts and channels; lock the screen; avoid personal sync"),
    ("External media", "Confirm transfer completion and safely eject before removal"),
], [1.5, 5.0])

chapter(16, "Print, scan and mail: procedural facts need a trace",
    "Printers and scanners often sit between evidence, delivery and a deadline.",
    bullet_items=(
        "Before printing, confirm color, sides, paper size, copies, pagination and binding; inspect the output page by page.",
        "Before scanning, remove fasteners and preserve order; after scanning, check count, orientation, clarity and missing pages.",
        "Keep a copy, dispatch record and tracking trail for mailed material.",
        "Do not dismantle equipment or destroy an uncertain document; follow the manual and escalate.",
    ))

section_break("Part IV  Professional delivery", "Make every result usable and reviewable", "part4")

chapter(17, "Stop document errors before submission",
    "A polished sentence cannot rescue the wrong party, version, attachment or deadline.",
    bullet_items=(
        "Confirm the recipient, purpose and approved template.",
        "Check names, identifiers, dates, numbers, facts, citations and defined terms.",
        "Remove comments and revision traces; verify signatures, seals and attachments.",
        "Open the final file from its delivery location, not only from the editor.",
        "Record who reviewed and what version was released.",
    ))

chapter(18, "Evidence organization is not numbering",
    "A numbered pile is not a fact structure.",
    paragraphs=("Preserve source identity, order and integrity. Build a timeline that links date, event, person, supporting material and unresolved question. Separate what the material proves from what someone hopes it proves.",),
    bullet_items=(
        "Keep originals and working copies distinct.",
        "Record missing pages, poor scans and conflicting dates instead of silently repairing them.",
        "Use an index that lets another person locate each item quickly.",
        "Check a sample before applying page numbers or names to the entire set.",
    ))

chapter(19, "A file should say honestly what is missing",
    "Completeness does not mean manufacturing a document that never existed.",
    paragraphs=("Choose the structure from the matter type, role and current stage. Mark absent, not yet received and not applicable as different states. A future reviewer must be able to see both the record and its gaps.",))

chapter(20, "Group first, review second",
    "When accepting or submitting material, classification is a risk-control step.",
    numbered_items=(
        "Identify the matter, parties, role, stage and deadline.",
        "Group authority documents, claims, evidence, originals and administrative forms.",
        "Count and record originals before marking, binding or transferring them.",
        "Match each scan to the physical document.",
        "Have the accountable professional confirm the final list and number of copies.",
    ))

chapter(21, "If information never enters the system, it was not retained",
    "Memory is not a shared record.",
    paragraphs=("Record the task, owner, timing, source, current version, result, location and next action in the team's approved system. The objective is not bureaucratic reporting; it is continuity when the original person is absent.",),
    quote_text="A useful record lets another person continue the work without reconstructing private memory.")

chapter(22, "Professionalism hides in the physical scene",
    "Visitors should not have to ask whether the team is prepared.",
    bullet_items=(
        "Prepare the room, materials, privacy and basic hospitality before a meeting.",
        "Use and reset the equipment the team actually has - for example printing, conferencing, display or beverage equipment.",
        "After the meeting, secure sensitive material, record agreed actions and restore the shared space.",
        "For photography or recording, confirm authorization, battery, storage, framing, sound, privacy and filing before release.",
    ))

chapter(23, "The stronger AI becomes, the more clearly a human must own the result",
    "AI makes average first drafts cheap. Problem definition, judgment, taste and responsibility become more valuable.",
    paragraphs=(
        "Do not begin with a tool list. Begin with the capability the task requires, the information boundary and the human acceptance test.",
        "A fluent answer is not a verified answer. A person defines the goal, supplies lawful context, tests sources and facts, decides what is fit for use and signs off on the result.",
    ))
add_table(["Stage", "Human responsibility", "Useful AI role"], [
    ("Define", "State the real problem, audience, risk and success test", "Suggest questions and alternative decompositions"),
    ("Prepare", "Select permitted, public, fictional or sufficiently de-identified material", "Organize input and identify missing context"),
    ("Generate", "Set constraints and request evidence or uncertainty", "Draft, compare, summarize, simulate an opposing view"),
    ("Verify", "Check sources, facts, law, numbers, confidentiality and tone", "Highlight contradictions and run repeated checks"),
    ("Deliver", "Approve, take responsibility and store the verified version", "Format and automate repeatable steps"),
], [1.0, 3.45, 2.05])
doc.add_heading("AI red lines", level=2)
bullets([
    "Do not enter client secrets, identity data, evidence, contacts or internal strategy without explicit authorization.",
    "Do not invent a source, quotation, authority, event or document.",
    "Do not deliver AI output before a qualified human has checked the facts and the intended use.",
    "Do not let the tool decide what is worth doing or who bears the consequence.",
])
quote("AI can accelerate your work. It cannot decide what matters or what is correct.", "Human agency")

section_break("Part V  Direction", "Turn an internship into portable capability", "part5")

chapter(24, "Stand on a common foundation before choosing a direction",
    "Different careers still depend on the same early evidence of reliability.",
    bullet_items=(
        "Protect confidential and personal information.",
        "Understand the purpose and restate the task.",
        "Work from sources, facts and observable results.",
        "Make a sample, accept feedback and improve the second version.",
        "Name and store work so another person can use it.",
        "Report delay, conflict and uncertainty early.",
    ))

chapter(25, "Different majors, different entry points, one development logic",
    "A major is a starting hypothesis, not a permanent allocation of work.")
doc.add_heading("Law and related disciplines", level=2)
p("Legal knowledge is useful, but interns often overestimate memory of rules and underestimate fact organization. Start with timelines, evidence indexes, source-based research and short document drafts that can be checked.")
doc.add_heading("Visual, media and content disciplines", level=2)
p("Attractive legal communication that misleads is worse than unattractive communication. Start with information hierarchy, fact-checking, privacy, authorization, captions, editing and release workflows.")
doc.add_heading("Other or mixed backgrounds", level=2)
p("Identify a real problem the background can help solve, define the boundary, start with a small sample and use a concrete output as evidence. Do not force every person into an old route merely because the team already knows it.")

chapter(26, "Leave the result behind and take the capability with you",
    "An internship becomes durable when the team keeps a useful asset and the intern keeps a better method.",
    bullet_items=(
        "Write a report that records tasks, errors, capability change, career judgment and next steps.",
        "Improve at least one checklist, template, SOP or knowledge entry.",
        "Complete one output that can be reviewed: research note, document draft, evidence structure, design or content work.",
        "Handover unfinished work, files, accounts, equipment and confidentiality obligations.",
        "Ask what kind of work you now understand better - and what evidence still needs to be gathered before choosing a career direction.",
    ), quote_text="A certificate proves that you were present. Your working method is what travels with you.")

section_break("Appendices", "Printable and reusable tools", "appendix")

doc.add_heading("Appendix 1  Onboarding and offboarding checklist", level=1)
add_table(["When", "Action", "Check"], [
    ("Before", "Confirm schedule, insurance or required forms, work location and known conditions.", "□"),
    ("Arrival", "Sign confidentiality and necessary acknowledgements; verify identity or student status only as required.", "□"),
    ("Accounts", "Receive approved work account, computer, phone, key or access and record custody.", "□"),
    ("Day 1", "Learn information boundaries, storage, question order, urgent reporting and shared space.", "□"),
    ("Departure", "Complete handover, return property, remove personal access and clear work data from personal devices.", "□"),
], [1.0, 4.95, 0.6])

doc.add_heading("Appendix 2  First-week baseline assessment", level=1)
add_table(["Area", "Passing evidence", "Result"], [
    ("Handbook", "Can locate the major rules and the correct question path", "□ pass  □ practise"),
    ("Computer and storage", "Can find, name, save, move and file work", "□ pass  □ practise"),
    ("PDF and equipment", "Can scan, OCR, combine, print and handle basic faults", "□ pass  □ practise"),
    ("Confidentiality", "Understands storage, transmission, social and AI boundaries", "□ pass  □ practise"),
    ("Communication", "Can restate a task, confirm timing and report progress or risk", "□ pass  □ practise"),
    ("Quality", "Names, pages, attachments and locations are basically accurate", "□ pass  □ practise"),
], [1.25, 4.05, 1.25])

doc.add_heading("Work-computer acceptance tasks", level=2)
add_table(["Task", "Passing standard", "Result"], [
    ("Windows and keyboard", "Switch, close, quit, lock and use the required shortcuts", "□ pass  □ practise"),
    ("Files", "Create, sort, search, rename, copy, move and confirm the path", "□ pass  □ practise"),
    ("PDF", "Create, combine, split, rotate, OCR and reopen a checked result", "□ pass  □ practise"),
    ("Shared storage", "Access the approved directory, preserve the source and verify the saved copy", "□ pass  □ practise"),
    ("Security", "Avoid personal sync and lock the screen when leaving", "□ pass  □ practise"),
], [1.2, 4.25, 1.1])

doc.add_heading("Appendix 3  Task delivery card", level=1)
for label in ["Task", "Owner", "Purpose", "Output", "Reference or file location", "Deadline", "Review point", "Main risks", "Completion and storage location"]:
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(7)
    font(para.add_run(label + ": "), 10.5, True, NAVY)
    font(para.add_run("____________________________________________________________"), 10.5, color=MUTED)

doc.add_heading("Appendix 4  Weekly review", level=1)
for question in [
    "What did I complete best this week, and why?",
    "What error or inefficiency appeared, and what was its root cause?",
    "What reusable method did I learn?",
    "What template, checklist or knowledge entry did I leave behind?",
    "What more difficult task should I try next?",
    "What did this week teach me about my direction?",
]:
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(5)
    font(para.add_run(question), 10.5, True, NAVY)
    for _ in range(2):
        p("________________________________________________________________________________", after=2)

doc.add_heading("Appendix 5  Knowledge-base entry", level=1)
add_table(["Field", "What to record"], [
    ("Title", "The problem this entry solves"),
    ("Use", "When to use it and when not to"),
    ("Prerequisites", "Accounts, material, permission or equipment"),
    ("Steps", "Ordered actions, with screenshots when useful"),
    ("Common failures", "What is easy to miss and what risk follows"),
    ("Location", "Shared path, template name or related record"),
    ("Maintenance", "Author, reviewer and last update"),
], [1.3, 5.25])

doc.add_heading("Appendix 6  Internship outputs", level=1)
bullets([
    "Personal internship report: work, capability change, errors, career judgment and next steps.",
    "One reviewed communication output, public or internal, with facts, privacy and attribution checked.",
    "At least one improvement to a handbook, workflow, checklist, template or knowledge entry.",
    "One handover covering unfinished tasks, file locations, accounts, equipment and confidential material.",
])

doc.add_heading("Appendix 7  Release-gate SOP template", level=1)
add_table(["Module", "What to record"], [
    ("Matter", "Name, date, operator, reviewer and final user"),
    ("Outcome", "The required final state and objective acceptance test"),
    ("Start conditions", "Permissions, materials, equipment, version, time and people"),
    ("Red lines", "Any condition that must stop the work"),
    ("Small sample", "One page, file, set or trial accepted before scaling"),
    ("Execution", "Critical steps, checks, exceptions and response"),
    ("External release", "Recipient, final version, attachments, authorization and owner approval"),
    ("Close", "Filing, backup, originals, equipment, space and handover"),
    ("Log", "Start/end, quantity, location, exception, remedy and improvement"),
    ("Status", "□ release  □ hold for correction  □ cancel"),
], [1.3, 5.25])

doc.add_heading("Appendix 8  Manager's assignment and coaching check", level=1)
add_table(["Stage", "Manager check", "Done"], [
    ("Before", "Purpose, output, standard, sample, permission, timing, risk and reviewer are clear.", "□"),
    ("Understanding", "The intern restated the task and conflicting instructions were resolved.", "□"),
    ("Sample", "A small sample was accepted before batch or high-risk work expanded.", "□"),
    ("Process", "Review points exist and urgent-report situations are known.", "□"),
    ("Result", "Accept or return is explicit, with observable reasons and next action.", "□"),
    ("Development", "Reliable basic work leads to harder, more valuable work.", "□"),
    ("Knowledge", "A new lesson or exception entered the checklist or knowledge base.", "□"),
    ("Authority", "The standard concerns work, not status, fear or personality control.", "□"),
    ("Self-review", "Did vague assignment, multiple owners or delayed feedback create rework?", "□"),
], [1.15, 4.85, 0.55])

doc.add_heading("Appendix 9  Newcomer calibration meeting", level=1)
add_table(["Time", "Agenda", "Facilitator cue"], [
    ("3 min", "State the purpose and the concrete collaboration problem", "Use facts, not labels"),
    ("6 min", "Short introductions and roles", "30-60 seconds each"),
    ("8 min", "Purpose, first-week standard and staged tasks", "Name the time and evidence"),
    ("9 min", "Each mentor states three core requirements", "Resolve conflicts in the room"),
    ("6 min", "Interns state needs, questions and expectations", "Do not interrupt simple questions"),
    ("8 min", "Confirm next task, owner, sample and deadline", "Create a written action list"),
], [0.95, 3.1, 2.45])

doc.add_heading("Appendix 10  Reading index from previous interns", level=1)
p("The linked reflections were published in Chinese on the WeChat account 没事多吃核桃. Read them as evidence from different starting points, not as scripts every future intern must imitate.")
for title_text, url, note in [
    ("First time observing a hearing", "https://mp.weixin.qq.com/s/9cRjuC24GyRrhRYt2_wHug", "Compare real procedure with imagined courtroom work."),
    ("First impressions of a law-office internship", "https://mp.weixin.qq.com/s/70rtliJ2Ntfo2tlskVeUeA", "Professional work, collaboration, time and ethics."),
    ("Learning without end", "https://mp.weixin.qq.com/s/6D6WGX1Hk5_wAzrWVeDb4A", "Basic work, questions, review and execution."),
    ("From page-number errors to independent filing", "https://mp.weixin.qq.com/s/ocqJtVvjPOtdGo0KvmsFmQ", "How specific feedback changes capability."),
    ("What can I contribute to a team?", "https://mp.weixin.qq.com/s/AAk6JgKd9kl8OCzx07Z1uQ", "Platform, respect and the use of growth time."),
    ("Editing and design inside legal practice", "https://mp.weixin.qq.com/s/OHRfieulR5-LFk8dtkFJTA", "Information design without losing professional rigor."),
    ("A small fish's adventure called growth", "https://mp.weixin.qq.com/s/x_q8x0s_fh9pSj81h1fJyQ", "Self-learning, equipment, verification and SOP thinking."),
]:
    linked_bullet(title_text, url, note)

doc.add_heading("Appendix 11  Recruitment-to-onboarding handoff", level=1)
p("This card does not label a candidate. It turns verified recruitment facts into a first-week plan and leaves room for the newcomer to correct the record.")
add_table(["Field", "Record only", "Convert into"], [
    ("Need", "Why the team is recruiting and the real problem to solve", "One primary task line"),
    ("Conditions", "Dates, continuity, location, schedule and confirmed compensation boundaries", "Realistic workload and unresolved items"),
    ("Starting evidence", "Relevant courses, work, projects, tools and transferable skills", "A first small task that uses a strength"),
    ("Work sample", "Task, time, AI use, observed result, errors and revision", "First-week checks and reduced retesting where justified"),
    ("Learning agility", "Can the person restate, apply, question and verify a new rule?", "Depth of demonstration and review"),
    ("Feedback", "Response to specific feedback and quality of revision", "Communication frequency and sample checks"),
    ("Expectations", "What the candidate hopes to learn or test", "Early response to what can and cannot be provided"),
    ("First-week plan", "Foundation, strength, practice need, first output and owner", "Task card, sample, deadline and reviewer"),
    ("Risk", "Confidentiality, conflict, permission, time and urgent reports", "Complete before opening higher-risk access"),
], [1.15, 3.1, 2.4])
bullets([
    "Use short samples that do not contain real confidential material or unpaid production work.",
    "Record behavior and revision, not labels such as smart, weak or strange.",
    "Restrict access, state retention and remove copies that no longer need to exist.",
])

doc.add_heading("Conclusion  Do not drift through the opportunity", level=1)
p("The easiest thing to waste when young is not a single day. It is the habit of doing only what cannot be avoided.", style="Lead")
p("You do not need a perfect career answer on day one. Make the next small task trustworthy, then ask where the method is taking you. Accept experience and rules without surrendering independent judgment.")
quote("A certificate proves that you were there. Your way of working is what stays with you.", "For every intern")

doc.add_heading("Sources and further reading", level=2)
p("The guide draws on the growth-oriented style of the Shanghai Jiao Tong University Survival Guide; checklist and inversion thinking associated with Poor Charlie's Almanack; research on work engagement and newcomer socialization; scholarship on paternalistic leadership and its boundaries; and Qiu Yiwu's public AI course materials. These ideas are adapted through practice rather than quoted as a substitute for judgment.")
sources = doc.add_paragraph()
add_hyperlink(sources, "Shanghai Jiao Tong University Survival Guide", "https://survivesjtu.gitbook.io/survivesjtumanual")
font(sources.add_run(" | "), 10.5, color=MUTED)
add_hyperlink(sources, "Poor Charlie's Almanack", "https://press.stripe.com/poor-charlies-almanack")
font(sources.add_run(" | "), 10.5, color=MUTED)
add_hyperlink(sources, "AI in Business", "https://qiuyiwu.com/zhongda/")
font(sources.add_run(" | "), 10.5, color=MUTED)
add_hyperlink(sources, "AI Thinking and Cognition", "https://qiuyiwu.com/ai-cognition/")

doc.core_properties.title = "The Intern Survival Guide - Public Edition"
doc.core_properties.subject = "Intern development, professional work and team management"
doc.core_properties.author = "George Hu (Hu Zexian)"
doc.core_properties.keywords = "internship, onboarding, mentoring, confidentiality, SOP, checklist, AI, professional development"
doc.core_properties.comments = "English public reading edition of The Intern Survival Guide."
doc.core_properties.category = "Training and management"
doc.core_properties.language = "en-US"
doc.core_properties.last_modified_by = "George Hu"
doc.core_properties.created = datetime(2026, 7, 16, 8, 0, 0)
doc.core_properties.modified = datetime(2026, 7, 16, 8, 0, 0)
doc.core_properties.version = "2026.07-en"

visible = "\n".join(
    [paragraph.text for paragraph in doc.paragraphs]
    + [cell.text for table in doc.tables for row in table.rows for cell in row.cells]
)
for forbidden in ("NAS", "Workbench", "HU ZE XIAN", "内部学习资料"):
    if forbidden in visible:
        raise ValueError(f"English public edition still contains {forbidden!r}")
for required in ("The Intern Survival Guide", "Chapter 26", "Appendix 11", "George Hu"):
    if required not in visible:
        raise ValueError(f"English public edition is missing {required!r}")

doc.save(str(OUT))
print(OUT)
