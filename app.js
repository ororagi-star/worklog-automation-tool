const form = document.querySelector("#worklog-form");
const workDate = document.querySelector("#work-date");
const author = document.querySelector("#author");
const done = document.querySelector("#done");
const issues = document.querySelector("#issues");
const next = document.querySelector("#next");
const result = document.querySelector("#result");
const copyButton = document.querySelector("#copy-button");
const clearButton = document.querySelector("#clear-button");

const today = new Date();
workDate.value = today.toISOString().slice(0, 10);

function linesToBullets(text, fallback) {
  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length === 0) {
    return `- ${fallback}`;
  }

  return lines.map((line) => `- ${line}`).join("\n");
}

function buildWorklog() {
  const dateText = workDate.value || today.toISOString().slice(0, 10);
  const authorText = author.value.trim() || "미입력";

  return [
    `# 업무일지 - ${dateText}`,
    "",
    `작성자: ${authorText}`,
    "",
    "## 오늘 한 일",
    linesToBullets(done.value, "작성된 업무가 없습니다."),
    "",
    "## 이슈 / 막힌 점",
    linesToBullets(issues.value, "특이사항 없음"),
    "",
    "## 내일 할 일",
    linesToBullets(next.value, "작성된 예정 업무가 없습니다."),
  ].join("\n");
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  result.textContent = buildWorklog();
});

copyButton.addEventListener("click", async () => {
  await navigator.clipboard.writeText(result.textContent);
  copyButton.textContent = "복사됨";

  window.setTimeout(() => {
    copyButton.textContent = "복사";
  }, 1200);
});

clearButton.addEventListener("click", () => {
  done.value = "";
  issues.value = "";
  next.value = "";
  result.textContent = "왼쪽에 내용을 입력하면 업무일지가 만들어집니다.";
});
