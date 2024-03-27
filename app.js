const $ = selector => document.querySelector(selector);
const listen = (host, type, handler) => host.addEventListener(type, handler);

window.onload = function () {
    let spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));

    const importPassword = $('#importPassword');
    const selectBtn = $('#selectBtn');
    const fileSelect = $('#selectedFile');
    const importBtn = $('#importBtn');
    const protectBtn = $('#protectBtn');
    const warningBox = $('#warningBox');
    const exportPassword = $('#exportPassword');
    const exportBtn = $('#exportBtn');

    listen(selectBtn, "click", () => fileSelect.click());

    const fileSelectedHandler = () => {
        importPassword.disabled = false;
        importBtn.disabled = false;
    }

    listen(fileSelect, 'change', fileSelectedHandler);

    const wrongPasswordHandler = message => {
        importPassword.setAttribute('warning', true);
        importPassword.style.animation = "shake 0.5s";
        setTimeout(() => importPassword.style.animation = "", 500);
        warningBox.innerText = message;
        importPassword.value = '';
    };

    listen(importPassword, 'focus', () => {
        warningBox.innerText = '';
        importPassword.removeAttribute('warning');
    });

    const protectHandler = () => {
        var option = {
            allowSelectLockedCells:true,
            allowSelectUnlockedCells:true,
            allowFilter: true,
            allowSort: false,
            allowResizeRows: true,
            allowResizeColumns: false,
            allowEditObjects: false,
            allowDragInsertRows: false,
            allowDragInsertColumns: false,
            allowInsertRows: false,
            allowInsertColumns: false,
            allowDeleteRows: false,
            allowDeleteColumns: false,
            allowOutlineColumns: false,
            allowOutlineRows: false
        };
        spread.getSheet(0).options.protectionOptions = option;
        spread.getSheet(0).options.isProtected = true;
    };
    listen(protectBtn, 'click', protectHandler);

    const importFileHandler = () => {
        let file = fileSelect.files[0];
        if (!file) return ;
        spread.import(file, console.log, error => {
            if (error.errorCode === GC.Spread.Sheets.IO.ErrorCode.noPassword || error.errorCode === GC.Spread.Sheets.IO.ErrorCode.invalidPassword) {
                wrongPasswordHandler(error.errorMessage);
            }
        }, {
            fileType: GC.Spread.Sheets.FileType.excel,
            password: importPassword.value
        });
    };
    listen(importBtn, 'click', importFileHandler);

    const exportFileHandler = () => {
        let password = exportPassword.value;
        spread.export(blob => saveAs(blob, (password ? 'encrypted-' : '') + 'export.xlsx'), console.log, {
            fileType: GC.Spread.Sheets.FileType.excel,
            password: password
        });
    };
    listen(exportBtn, 'click', exportFileHandler);
};