// Study Quest's reflection and goal export. Called with the current attempt's saved data.
function buildReflectionsAndGoalsText(progress, goals, attemptNumber) {
  const entries = sections.filter(section => {
    const saved = progress[section.id] || {};
    return (saved.reflectionCompleted && saved.reflectionText && saved.reflectionText.trim()) ||
      (goals[section.id] && goals[section.id].trim());
  });
  if (!entries.length) return "";

  const lines = [
    "STUDY QUEST: MY REFLECTIONS AND GOALS",
    "=".repeat(46),
    `Downloaded: ${new Date().toLocaleDateString()}`,
    `Attempt: ${attemptNumber}`,
    "The reflection prompts invite a response to one or more questions; each saved answer appears below.",
    ""
  ];
  for (const section of entries) {
    const saved = progress[section.id] || {};
    const realm = getRealm(section.realm);
    lines.push(`${realm ? realm.title : section.realm} / ${section.diagnosticTitle} - ${section.fullTitle}`);
    lines.push("-".repeat(46));
    lines.push("Reflection prompts:");
    getReflectionPrompts(attemptNumber).forEach((prompt, index) => lines.push(`${index + 1}. ${prompt}`));
    lines.push("My reflection:");
    lines.push(saved.reflectionCompleted && saved.reflectionText && saved.reflectionText.trim()
      ? saved.reflectionText.trim() : "No reflection saved yet.");
    lines.push("My goal:");
    lines.push(goals[section.id] && goals[section.id].trim()
      ? `${goals[section.id].trim()}${saved.smartGoalSaved ? "" : " (draft; not marked complete)"}`
      : "No goal saved yet.");
    lines.push("");
  }
  return lines.join("\n");
}

function downloadReflectionsAndGoals(progress, goals, attemptNumber) {
  const content = buildReflectionsAndGoalsText(progress, goals, attemptNumber);
  if (!content) return;
  const url = URL.createObjectURL(new Blob(["\uFEFF", content], {type: "text/plain;charset=utf-8"}));
  const link = document.createElement("a");
  link.href = url;
  link.download = `study-quest-reflections-and-goals-attempt-${attemptNumber}.txt`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
