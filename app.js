const $ = (s) => document.querySelector(s);

const S = {
  grid: 3, allQs: [], qs: [], img: null, time: 300, left: 300,
  i: 0, done: 0, right: 0, open: new Set(), locked: new Set(),
  timer: null, end: false, student: false, name: "", answerMode: "hide"
};

document.querySelectorAll("#gridChoices button").forEach((b) => {
  b.addEventListener("click", () => {
    S.grid = Number(b.dataset.grid);
    document.querySelectorAll("#gridChoices button").forEach((x) => x.classList.toggle("active", x === b));
    checkReady();
  });
});
document.querySelector("#gridChoices button").classList.add("active");

$("#imageInput").addEventListener("change", (e) => {
  const f = e.target.files?.[0];
  if (!f) return;
  $("#imageName").textContent = f.name;
  const r = new FileReader();
  r.onload = () => {
    S.img = r.result;
    $("#preview").innerHTML = "<img src='" + r.result + "' alt='Ảnh xem trước'>";
    checkReady();
  };
  r.readAsDataURL(f);
});

$("#excelInput").addEventListener("change", async (e) => {
  const f = e.target.files?.[0];
  if (!f) return;
  $("#excelName").textContent = f.name;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet) throw new Error("Không có trang tính");
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    S.allQs = rows.map((r) => {
      const n = {};
      Object.keys(r).forEach((k) => n[String(k).trim().toLowerCase()] = r[k]);

      const get = (...keys) => {
        for (const key of keys) {
          const v = n[key];
          if (v !== undefined && v !== null && String(v).trim() !== "") return String(v).trim();
        }
        return "";
      };

      const rawAns = get("đáp án", "dap an", "answer");
      const ans = rawAns.toUpperCase().replace(/[^ABCD]/g, "").charAt(0);

      return {
        q: get("câu hỏi", "cau hoi", "question"),
        A: get("a", "đáp án a"),
        B: get("b", "đáp án b"),
        C: get("c", "đáp án c"),
        D: get("d", "đáp án d"),
        ans
      };
    }).filter((x) => x.q && x.A && x.B && x.C && x.D && "ABCD".includes(x.ans));

    checkReady();
  } catch (err) {
    S.allQs = [];
    $("#validation").textContent = "Không đọc được file Excel/CSV. Hãy kiểm tra đúng mẫu cột: Câu hỏi, A, B, C, D, Đáp án.";
    $("#validation").className = "validation err";
    $("#startBtn").disabled = true;
    $("#shareBtn").disabled = true;
  }
});

function checkReady() {
  const needed = S.grid * S.grid;
  const ok = Boolean(S.img) && S.allQs.length >= needed;
  $("#validation").textContent = ok
    ? "✓ Sẵn sàng: " + needed + " mảnh / " + needed + " câu hỏi"
    : "Cần ảnh và ít nhất " + needed + " câu hỏi hợp lệ.";
  $("#validation").className = ok ? "validation ok" : "validation err";
  $("#startBtn").disabled = !ok;
  $("#shareBtn").disabled = !ok;
}

$("#time").addEventListener("change", (e) => S.time = Number(e.target.value));
$("#answerMode").addEventListener("change", (e) => S.answerMode = e.target.value);
$("#startBtn").addEventListener("click", () => startGame(false));
$("#shareBtn").addEventListener("click", createShareLink);
$("#copyShareBtn").addEventListener("click", copyShareLink);
$("#studentStart").addEventListener("click", () => {
  S.student = true;
  S.name = $("#studentName").value.trim() || "Học sinh";
  startGame(true);
});
$("#finishBtn").addEventListener("click", finishGame);
$("#newGameBtn").addEventListener("click", () => location.href = location.pathname);
$("#resetBtn").addEventListener("click", () => location.href = location.pathname);

function prepareQuestions() {
  const needed = S.grid * S.grid;
  S.left = S.time;
  S.done = 0; S.right = 0; S.i = 0; S.end = false;
  S.open = new Set(); S.locked = new Set();

  S.qs = S.allQs.slice(0, needed);
  if (!S.student && $("#order").value === "shuffle") {
    S.qs = [...S.qs].sort(() => Math.random() - 0.5);
  }
}

function startGame(studentMode) {
  S.student = Boolean(studentMode);
  prepareQuestions();

  $("#setup").classList.add("hidden");
  $("#student").classList.add("hidden");
  $("#result").classList.add("hidden");
  $("#game").classList.remove("hidden");

  $("#gameMeta").textContent = S.grid + "×" + S.grid + " • " + S.qs.length + " câu" + (S.student ? " • " + S.name : "");
  $("#timer").textContent = formatTime(S.left);

  const board = $("#board");
  board.style.gridTemplateColumns = "repeat(" + S.grid + ", minmax(0, 1fr))";
  board.innerHTML = "";

  S.qs.forEach((_, index) => {
    const tile = document.createElement("div");
    tile.className = "tile";
    tile.innerHTML = "<span class='num'>" + (index + 1) + "</span><img alt='Mảnh ghép " + (index + 1) + "'>";
    const img = tile.querySelector("img");
    const row = Math.floor(index / S.grid);
    const col = index % S.grid;
    img.style.width = (S.grid * 100) + "%";
    img.style.height = (S.grid * 100) + "%";
    img.style.left = (-col * 100) + "%";
    img.style.top = (-row * 100) + "%";
    board.appendChild(tile);
  });

  updateUI();
  showQuestion();
  clearInterval(S.timer);
  S.timer = setInterval(tick, 1000);
}

function showQuestion() {
  if (S.end) return;
  if (S.done >= S.qs.length) {
    finishGame();
    return;
  }

  S.i = S.done;
  const q = S.qs[S.i];
  $("#qIndex").textContent = "Câu " + (S.i + 1) + " / " + S.qs.length;
  $("#qStatus").textContent = "Chưa trả lời";
  $("#question").textContent = q.q;
  $("#feedback").textContent = "";
  $("#feedback").className = "feedback";
  $("#answers").innerHTML = "";

  ["A", "B", "C", "D"].forEach((key) => {
    const btn = document.createElement("button");
    btn.className = "answer";
    btn.innerHTML = "<b>" + key + ".</b><span>" + escapeHtml(q[key]) + "</span>";
    btn.addEventListener("click", () => answerQuestion(key));
    $("#answers").appendChild(btn);
  });
}

function answerQuestion(key) {
  if (S.end) return;

  const q = S.qs[S.i];
  document.querySelectorAll(".answer").forEach((b) => b.disabled = true);
  S.done++;

  const correct = key === q.ans;
  $("#qStatus").textContent = correct ? "Đúng" : "Sai";

  if (correct) {
    S.right++;
    S.open.add(S.i);
    const tile = document.querySelectorAll(".tile")[S.i];
    tile.classList.add("open");
    tile.querySelector("img").src = S.img;
    $("#feedback").textContent = "✓ Chính xác! Mảnh ghép đã mở.";
    $("#feedback").className = "feedback ok";
  } else {
    S.locked.add(S.i);
    document.querySelectorAll(".tile")[S.i].classList.add("locked");
    $("#feedback").textContent = "✕ Sai. Mảnh ghép không được mở; câu này đã khóa.";
    if (S.answerMode === "show") $("#feedback").textContent += " Đáp án: " + q.ans;
    $("#feedback").className = "feedback bad";
  }

  updateUI();
  setTimeout(showQuestion, 650);
}

function updateUI() {
  $("#score").textContent = (S.right * 100) + " điểm";
  $("#progressText").textContent = S.done + "/" + S.qs.length + " câu • Mở " + S.open.size + " mảnh";
  $("#progressBar").style.width = (S.done / S.qs.length * 100) + "%";
}

function tick() {
  if (S.end) return;
  S.left = Math.max(0, S.left - 1);
  $("#timer").textContent = formatTime(S.left);
  if (S.left === 0) finishGame();
}

function formatTime(sec) {
  return String(Math.floor(sec / 60)).padStart(2, "0") + ":" + String(sec % 60).padStart(2, "0");
}

function finishGame() {
  if (S.end) return;
  S.end = true;
  clearInterval(S.timer);

  $("#game").classList.add("hidden");
  $("#result").classList.remove("hidden");
  $("#resultSummary").innerHTML =
    (S.student ? "<b>" + escapeHtml(S.name) + "</b> — " : "") +
    "Đúng <b>" + S.right + "/" + S.qs.length + "</b> câu · mở được <b>" +
    S.open.size + "</b> mảnh · <b>" + (S.right * 100) + " điểm</b>.";

  const resultBoard = $("#resultBoard");
  resultBoard.style.gridTemplateColumns = "repeat(" + S.grid + ", minmax(0, 1fr))";
  resultBoard.innerHTML = "";

  S.qs.forEach((_, index) => {
    const tile = document.createElement("div");
    tile.className = "tile " + (S.open.has(index) ? "open" : "locked");
    tile.innerHTML = "<span class='num'>" + (index + 1) + "</span><img alt='Mảnh ghép " + (index + 1) + "'>";
    const img = tile.querySelector("img");
    const row = Math.floor(index / S.grid);
    const col = index % S.grid;
    img.style.width = (S.grid * 100) + "%";
    img.style.height = (S.grid * 100) + "%";
    img.style.left = (-col * 100) + "%";
    img.style.top = (-row * 100) + "%";
    if (S.open.has(index)) img.src = S.img;
    resultBoard.appendChild(tile);
  });
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (c) => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;" }[c]));
}

function encodeData(obj) {
  const bytes = new TextEncoder().encode(JSON.stringify(obj));
  let binary = "";
  bytes.forEach((b) => binary += String.fromCharCode(b));
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function decodeData(text) {
  const base = text.replace(/-/g, "+").replace(/_/g, "/") + "=".repeat((4 - text.length % 4) % 4);
  const binary = atob(base);
  const bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

function resizeImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = 800; canvas.height = 450;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, 800, 450);
      resolve(canvas.toDataURL("image/jpeg", 0.55));
    };
    img.onerror = reject;
    img.src = dataUrl;
  });
}

async function createShareLink() {
  if (!S.img || S.allQs.length < S.grid * S.grid) return;

  const button = $("#shareBtn");
  button.disabled = true;
  button.textContent = "⏳ Đang tạo link...";

  try {
    const img = await resizeImage(S.img);
    const needed = S.grid * S.grid;
    const data = {
      v: 2,
      title: "Lật mảnh ghép",
      grid: S.grid,
      time: S.time,
      answerMode: S.answerMode,
      qs: S.allQs.slice(0, needed),
      img
    };
    const link = location.origin + location.pathname + "#play=" + encodeData(data);
    $("#shareLink").value = link;
    $("#shareBox").classList.remove("hidden");
    $("#shareBox small").textContent = link.length > 180000
      ? "⚠ Link dài. Nếu gửi không được, cần chuyển sang chế độ phòng trực tuyến có máy chủ."
      : "Học sinh mở link trên điện thoại/iPad/máy tính và tự làm bài. Không cần đăng nhập.";

    if (navigator.clipboard && navigator.clipboard.writeText) {
      try { await navigator.clipboard.writeText(link); } catch (_) {}
    }
  } catch (err) {
    alert("Không tạo được link giao bài. Hãy thử ảnh nhỏ hơn.");
  } finally {
    button.disabled = false;
    button.textContent = "🔗 Giao bài cho học sinh";
  }
}

async function copyShareLink() {
  const value = $("#shareLink").value;
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(value);
    } else {
      $("#shareLink").focus(); $("#shareLink").select(); document.execCommand("copy");
    }
    $("#copyShareBtn").textContent = "✓ Đã sao chép";
  } catch (_) {
    $("#shareLink").focus(); $("#shareLink").select();
  }
  setTimeout(() => $("#copyShareBtn").textContent = "Sao chép", 1800);
}

function loadStudentLink() {
  if (!location.hash.startsWith("#play=")) return;

  try {
    const d = decodeData(location.hash.slice(6));
    if (!d || d.v !== 2 || !Array.isArray(d.qs) || !d.img) throw new Error("invalid");

    S.student = true;
    S.grid = Number(d.grid);
    S.time = Number(d.time);
    S.answerMode = d.answerMode || "hide";
    S.allQs = d.qs;
    S.img = d.img;

    $("#setup").classList.add("hidden");
    $("#student").classList.remove("hidden");
    $("#studentTitle").textContent = d.title || "Lật mảnh ghép";
    $("#studentMeta").textContent = S.grid + "×" + S.grid + " • " + d.qs.length + " câu • " + Math.round(S.time / 60) + " phút";
  } catch (err) {
    alert("Link giao bài không hợp lệ hoặc đã bị cắt.");
  }
}

checkReady();
loadStudentLink();