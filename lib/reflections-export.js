// Create a Word document from the current attempt's locally saved work.
function buildReflectionsAndGoalsDocument(progress, goals, attemptNumber, sectionId) {
  const entries = sections.filter(section => {
    if (sectionId && section.id !== sectionId) return false;
    const saved = progress[section.id] || {};
    return (saved.reflectionCompleted && saved.reflectionText && saved.reflectionText.trim()) ||
      (goals[section.id] && goals[section.id].trim());
  });
  if (!entries.length) return null;

  const {Document, Paragraph, TextRun, HeadingLevel} = docx;
  const paragraph = (value, options = {}) => new Paragraph({
    children: String(value).split(/\r\n|\r|\n/).map((line, index) =>
      new TextRun({text: line, break: index ? 1 : 0,
        ...(options.heading || options.style === "Title" ? {} : {font: "Arial", size: 22, color: "000000"})})),
    spacing: {after: 160, line: 300}, ...options
  });
  const children = [
    paragraph(sectionId ? "Study Quest Reflection and Goal" : "Study Quest Reflections and Goals", {style: "Title", spacing: {after: 240}}),
    paragraph(`Downloaded ${new Date().toLocaleDateString()}  |  Attempt ${attemptNumber}`),
    paragraph(sectionId
      ? "My saved reflection and goal for this Quest. The reflection responds to one or more of the prompts shown below."
      : "My saved reflections and goals from this Study Quest attempt. Each reflection responds to one or more of the prompts shown below.", {spacing: {after: 320}})
  ];

  for (const section of entries) {
    const saved = progress[section.id] || {};
    const realm = getRealm(section.realm);
    children.push(paragraph(`${realm ? realm.title : section.realm}  |  ${section.diagnosticTitle}`, {
      heading: HeadingLevel.HEADING_1, keepNext: true, spacing: {before: 320, after: 80}
    }));
    children.push(paragraph(section.fullTitle, {spacing: {after: 180}}));
    children.push(paragraph("Reflection prompts", {heading: HeadingLevel.HEADING_2, keepNext: true}));
    getReflectionPrompts(attemptNumber).forEach((prompt, index) =>
      children.push(paragraph(`${index + 1}. ${prompt}`, {spacing: {after: 90}})));
    children.push(paragraph("My reflection", {heading: HeadingLevel.HEADING_2, keepNext: true}));
    children.push(paragraph(saved.reflectionCompleted && saved.reflectionText && saved.reflectionText.trim()
      ? saved.reflectionText.trim() : "No reflection saved yet."));
    children.push(paragraph("My goal", {heading: HeadingLevel.HEADING_2, keepNext: true}));
    children.push(paragraph(goals[section.id] && goals[section.id].trim()
      ? `${goals[section.id].trim()}${saved.smartGoalSaved ? "" : " (draft; not marked complete)"}`
      : "No goal saved yet.", {spacing: {after: 260}}));
  }

  return new Document({
    creator: "Study Quest",
    title: "Study Quest Reflections and Goals",
    styles: {
      default: {document: {run: {font: "Arial", size: 22, color: "000000"}}},
      paragraphStyles: [
        {id: "Title", name: "Title", basedOn: "Normal", quickFormat: true,
          run: {font: "Arial", size: 32, bold: true, color: "000000"}},
        {id: "Heading1", name: "Heading 1", basedOn: "Normal", quickFormat: true,
          run: {font: "Arial", size: 27, bold: true, color: "000000"}},
        {id: "Heading2", name: "Heading 2", basedOn: "Normal", quickFormat: true,
          run: {font: "Arial", size: 23, bold: true, color: "000000"}}
      ]
    },
    sections: [{properties: {
      page: {size: {width: 12240, height: 15840}, margin: {top: 1080, right: 1080, bottom: 1080, left: 1080}}
    }, children}]
  });
}

async function downloadReflectionsAndGoals(progress, goals, attemptNumber, sectionId) {
  const document = buildReflectionsAndGoalsDocument(progress, goals, attemptNumber, sectionId);
  if (!document) return;
  try {
    const blob = await docx.Packer.toBlob(document);
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    link.href = url;
    link.download = `study-quest-${sectionId || "all-quests"}-reflection-and-goal-attempt-${attemptNumber}.docx`;
    window.document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (error) {
    console.error("Could not create the reflections and goals document:", error);
    window.alert("The document could not be downloaded. Your responses are still saved in this browser. Please try again.");
  }
}
