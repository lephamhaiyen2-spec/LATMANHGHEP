window.addEventListener("error",function(e){var b=document.querySelector("#bootError");if(b){b.classList.remove("hidden");b.textContent="⚠️ Ứng dụng gặp lỗi: "+(e.message||"JavaScript")+" — hãy tải lại trang.";}});
window.addEventListener("unhandledrejection",function(){var b=document.querySelector("#bootError");if(b){b.classList.remove("hidden");b.textContent="⚠️ Không xử lý được dữ liệu. Hãy tải lại trang và thử lại.";}});
window.addEventListener("error", (event) => {
  const box = document.querySelector("#bootError");
  if (box) {
    box.classList.remove("hidden");
    box.textContent = "⚠️ Ứng dụng gặp lỗi khi khởi động: " + (event.message || "lỗi JavaScript") + ". Hãy tải lại trang.";
  }
});
window.addEventListener("unhandledrejection", (event) => {
  const box = document.querySelector("#bootError");
  if (box) {
    box.classList.remove("hidden");
    box.textContent = "⚠️ Ứng dụng gặp lỗi khi xử lý dữ liệu. Hãy tải lại trang và thử lại.";
  }
});

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
  if (typeof XLSX === "undefined") {
    $("#validation").textContent = "⚠️ Thư viện Excel chưa tải được. Hãy tải lại trang hoặc kiểm tra kết nối Internet.";
    $("#validation").className = "validation err";
    return;
  }
  const f = e.target.files?.[0];
  if (!f) return;
  $("#excelName").textContent = f.name;

  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet) throw new Error("Không có trang tính");

    // Đọc trực tiếp theo ma trận để không phụ thuộc cách Excel đặt tên/định dạng header.
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    if (!matrix.length) throw new Error("File không có dữ liệu");

    const headers = matrix[0].map(normalizeHeader);
    const findCol = (...names) => {
      const wanted = names.map(normalizeHeader);
      return headers.findIndex(h => wanted.includes(h));
    };

    const qCol = findCol("Câu hỏi", "Cau hoi", "Question");
    const aCol = findCol("Đáp án A", "A");
    const bCol = findCol("Đáp án B", "B");
    const cCol = findCol("Đáp án C", "C");
    const dCol = findCol("Đáp án D", "D");
    const ansCol = findCol("Đáp án đúng", "Đáp án", "Dap an dung", "Dap an", "Answer");

    if ([qCol,aCol,bCol,cCol,dCol,ansCol].some(x => x < 0)) {
      throw new Error("HEADER");
    }

    S.allQs = matrix.slice(1).map((row) => {
      const q = cleanCell(row[qCol]);
      const A = cleanCell(row[aCol]);
      const B = cleanCell(row[bCol]);
      const C = cleanCell(row[cCol]);
      const D = cleanCell(row[dCol]);
      const rawAns = cleanCell(row[ansCol]);

      // Ưu tiên A/B/C/D; nếu ô đáp án chứa luôn nội dung đáp án thì đối chiếu với 4 lựa chọn.
      let ans = normalizeAnswer(rawAns);
      if (!ans) ans = inferAnswerFromText(rawAns, {A,B,C,D});

      return { q, A, B, C, D, ans };
    }).filter(x => x.q && x.A && x.B && x.C && x.D && /^[ABCD]$/.test(x.ans));

    const needed = S.grid * S.grid;
    if (S.allQs.length) {
      $("#validation").textContent =
        "✓ Đã đọc " + S.allQs.length + " câu hỏi hợp lệ. " +
        "Câu 1: đáp án đúng = " + S.allQs[0].ans + ".";
      $("#validation").className = S.allQs.length >= needed ? "validation ok" : "validation err";
    }

    checkReady();
  } catch (err) {
    S.allQs = [];
    $("#validation").textContent =
      err.message === "HEADER"
        ? "Không nhận diện được tiêu đề cột. Hãy dùng đúng mẫu: Câu hỏi | Đáp án A | Đáp án B | Đáp án C | Đáp án D | Đáp án đúng."
        : "Không đọc được file Excel. Hãy kiểm tra lại file mẫu.";
    $("#validation").className = "validation err";
    $("#startBtn").disabled = true;
    $("#shareBtn").disabled = true;
  }
});

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function cleanCell(value) {
  return String(value ?? "").trim();
}

function normalizeAnswer(value) {
  const t = cleanCell(value).toUpperCase();
  if (!t) return "";
  if (/^[ABCD]$/.test(t)) return t;

  const noAccent = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const m = noAccent.match(/(?:DAP\s*AN|ANSWER|OPTION|LUACHON)?\s*[:\-]?\s*\(?([ABCD])\)?(?:[.)\s]|$)/);
  if (m) return m[1];

  // Accept forms such as "A.", "A)", "(A)", "Đáp án: A", "answer A".
  const first = noAccent.match(/^\(?([ABCD])\)?(?:[.)\s:]|$)/);
  return first ? first[1] : "";
}

function normalizeCompare(value) {
  return cleanCell(value)
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.\)\(:\-]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function inferAnswerFromText(raw, options) {
  const target = normalizeCompare(raw);
  if (!target) return "";

  for (const key of ["A","B","C","D"]) {
    if (target === normalizeCompare(options[key])) return key;
  }
  return "";
}

function checkReady() {
  const needed = S.grid * S.grid;
  const hasImage = Boolean(S.img);
  const hasQuestions = S.allQs.length >= needed;
  const ok = hasImage && hasQuestions;
  if (ok) {
    $("#validation").textContent = "✓ Sẵn sàng: " + needed + " mảnh / " + needed + " câu hỏi hợp lệ.";
    $("#validation").className = "validation ok";
  } else {
    const missing = [];
    if (!hasImage) missing.push("ảnh bí mật");
    if (!hasQuestions) missing.push("ít nhất " + needed + " câu hỏi hợp lệ");
    $("#validation").textContent = "Cần " + missing.join(" và ") + ".";
    $("#validation").className = "validation err";
  }
  $("#startBtn").disabled = !ok;
  $("#shareBtn").disabled = !ok;
}

$("#time").addEventListener("change", (e) => S.time = Number(e.target.value));
$("#answerMode").addEventListener("change", (e) => S.answerMode = e.target.value);
$("#startBtn").addEventListener("click", () => startGame(false));
$("#shareBtn").addEventListener("click", createShareLink);
$("#copyShareBtn").addEventListener("click", copyShareLink);
$("#downloadQRBtn").addEventListener("click", downloadQR);
$("#templateBtn").addEventListener("click", downloadTemplate);
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

  const correct = normalizeCompare(key) === normalizeCompare(q.ans);
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

function resizeImage(dataUrl, width = 180, height = 101, quality = 0.16) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, width, height);
      let out = canvas.toDataURL("image/webp", quality);
      if (!out.startsWith("data:image/webp")) out = canvas.toDataURL("image/jpeg", quality);
      resolve(out);
    };
    img.onerror = reject;
    img.src = dataUrl;
  });
}

function getShareQuestions() {
  const needed = S.grid * S.grid;
  let qs = S.allQs.slice(0, needed);
  if ($("#order").value === "shuffle") qs = [...qs].sort(() => Math.random() - 0.5);
  return qs;
}

async function createShareLink() {
  if (!S.img || S.allQs.length < S.grid * S.grid) return;

  const button = $("#shareBtn");
  button.disabled = true;
  button.textContent = "⏳ Đang tạo link...";
  $("#shareBox").classList.add("hidden");

  try {
    const img = await resizeImage(S.img, 180, 101, 0.16);
    const qs = getShareQuestions();
    const data = {
      v: 3,
      title: "Lật mảnh ghép",
      grid: S.grid,
      time: S.time,
      answerMode: S.answerMode,
      order: $("#order").value,
      qs,
      img
    };
    let link = location.origin + location.pathname + "#play=" + encodeData(data);
    if (link.length > 7000) {
      data.img = await resizeImage(S.img, 120, 68, 0.10);
      link = location.origin + location.pathname + "#play=" + encodeData(data);
    }

    if (link.length > 8000) {
      throw new Error("LINK_TOO_LONG");
    }

    $("#shareLink").value = link;
    const openLink = $("#shareOpenLink");
    if (openLink) { openLink.href = link; openLink.textContent = "Mở link giao bài trên thiết bị khác ↗"; }
    $("#shareBox").classList.remove("hidden");
    $("#shareNote").textContent =
      "✓ Đã tạo link. " + qs.length + " câu hỏi sẽ được giao đúng theo thứ tự đã chọn.";
    $("#shareLimit").textContent =
      "Lưu ý: đây là link tự chứa dữ liệu, không cần máy chủ. Học sinh làm bài trên thiết bị riêng nhưng kết quả chưa tự gửi về máy giáo viên.";

    renderQR(link);

    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await navigator.clipboard.writeText(link);
        $("#copyShareBtn").textContent = "✓ Đã sao chép";
        setTimeout(() => $("#copyShareBtn").textContent = "Sao chép", 1800);
      } catch (_) {}
    }
  } catch (err) {
    const msg = err.message === "LINK_TOO_LONG"
      ? "Link quá dài. Hãy dùng ảnh nhẹ hơn hoặc giảm số mảnh ghép."
      : "Không tạo được link giao bài. Hãy thử ảnh nhỏ hơn.";
    alert(msg);
  } finally {
    button.disabled = false;
    button.textContent = "🔗 Giao bài cho học sinh";
  }
}

function renderQR(link) {
  const box = $("#qrCode");
  box.innerHTML = "";
  if (typeof QRCode === "undefined") {
    box.textContent = "QR chưa tải được. Bạn vẫn có thể mở bằng link.";
    return;
  }
  if (link.length > 2800) {
    box.textContent = "Link còn dài để QR quét ổn định. Hãy thử ảnh đơn giản hơn hoặc 3×3.";
    return;
  }
  try {
    new QRCode(box, {text:link,width:190,height:190,correctLevel:QRCode.CorrectLevel.L});
  } catch (_) {
    box.textContent = "Không tạo được QR. Hãy dùng nút Mở link giao bài.";
  }
}

function downloadQR() {
  const canvas = $("#qrCode canvas");
  const img = $("#qrCode img");
  let src = canvas ? canvas.toDataURL("image/png") : (img ? img.src : "");
  if (!src) return;
  const a = document.createElement("a");
  a.href = src;
  a.download = "QR-lat-manh-ghep.png";
  a.click();
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
    if (!d || ![2,3].includes(Number(d.v)) || !Array.isArray(d.qs) || !d.img) throw new Error("invalid");
    if (![3,4,5,6].includes(Number(d.grid)) || d.qs.length !== Number(d.grid) * Number(d.grid)) throw new Error("invalid");

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

function downloadTemplate() {
  if (typeof XLSX === "undefined") {
    alert("Thư viện Excel chưa tải được. Hãy tải lại trang rồi thử lại.");
    return;
  }
  const rows = [
    ["ID", "Câu hỏi", "Đáp án A", "Đáp án B", "Đáp án C", "Đáp án D", "Đáp án đúng"],
    [1, "DNA được cấu tạo từ những loại nucleotide nào?", "A, T, G, X", "A, U, G, X", "A, T, U, G", "T, U, G, X", "A"],
    [2, "Đơn vị cấu tạo nên phân tử DNA là:", "Gene", "Nucleotide", "Amino acid", "Protein", "B"],
    [3, "Trong phân tử DNA, nucleotide A liên kết bổ sung với nucleotide nào?", "G", "X", "T", "U", "C"],
    [4, "Quá trình tổng hợp RNA dựa trên khuôn DNA được gọi là:", "Tái bản", "Phiên mã", "Dịch mã", "Đột biến", "B"],
    [5, "Sản phẩm trực tiếp của quá trình phiên mã là:", "DNA", "Protein", "RNA", "Amino acid", "C"],
    [6, "Quá trình tổng hợp protein dựa trên thông tin của mRNA được gọi là:", "Tái bản", "Phiên mã", "Dịch mã", "Nhân đôi", "C"],
    [7, "Một đoạn DNA có trình tự trên một mạch là A–T–G–X–A–X. Trình tự mạch bổ sung là:", "T–A–X–G–T–G", "A–T–G–X–A–X", "U–A–X–G–U–G", "G–X–T–A–G–T", "A"],
    [8, "Nếu trình tự nucleotide của một gene thay đổi, điều gì có thể xảy ra?", "Luôn làm cơ thể chết", "Có thể làm thay đổi protein và tính trạng", "Chắc chắn không ảnh hưởng", "Luôn làm tăng chiều cao", "B"],
    [9, "Đơn phân cấu tạo nên nucleic acid là", "amino acid.", "protein.", "nucleotide.", "glucose.", "C"]
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:6},{wch:52},{wch:28},{wch:28},{wch:28},{wch:28},{wch:14}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Câu hỏi");
  XLSX.writeFile(wb, "mau-cau-hoi-lat-manh-ghep.xlsx");
}
