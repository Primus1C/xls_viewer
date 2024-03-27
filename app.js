const $ = (selector) => document.querySelector(selector);
const listen = (host, type, handler) => host.addEventListener(type, handler);

listen(selectBtn, "click", () => fileSelect.click());

const fileSelectedHandler = () => {
    importPassword.disabled = false;
    importBtn.disabled = false;
};

listen(fileSelect, "change", fileSelectedHandler);

const wrongPasswordHandler = (message) => {
    importPassword.setAttribute("warning", true);
    importPassword.style.animation = "shake 0.5s";
    setTimeout(() => (importPassword.style.animation = ""), 500);
    warningBox.innerText = message;
    importPassword.value = "";
};

listen(importPassword, "focus", () => {
    warningBox.innerText = "";
    importPassword.removeAttribute("warning");
});

const importFileHandler = () => {
    let file = fileSelect.files[0];
    if (!file) return;
    spread.import(
        file,
        console.log,
        (error) => {
            if (
                error.errorCode === GC.Spread.Sheets.IO.ErrorCode.noPassword ||
                error.errorCode ===
                    GC.Spread.Sheets.IO.ErrorCode.invalidPassword
            ) {
                wrongPasswordHandler(error.errorMessage);
            }
        },
        {
            fileType: GC.Spread.Sheets.FileType.excel,
            password: importPassword.value,
        }
    );
};
listen(importBtn, "click", importFileHandler);

const exportFileHandler = () => {
    let password = exportPassword.value;
    spread.export(
        (blob) => saveAs(blob, (password ? "encrypted-" : "") + "export.xlsx"),
        console.log,
        {
            fileType: GC.Spread.Sheets.FileType.excel,
            password: password,
        }
    );
};
listen(exportBtn, "click", exportFileHandler);


window.onload = function () {
    const importPassword = $("#importPassword");
    const selectBtn = $("#selectBtn");
    const fileSelect = $("#selectedFile");
    const importBtn = $("#importBtn");
    const warningBox = $("#warningBox");
    const exportPassword = $("#exportPassword");
    const exportBtn = $("#exportBtn");
    let spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
};

