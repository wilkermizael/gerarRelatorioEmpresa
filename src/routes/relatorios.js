"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var express_1 = require("express");
var exceljs_1 = require("exceljs");
var pdf_lib_1 = require("pdf-lib");
var fs_1 = require("fs");
var path_1 = require("path");
var router = (0, express_1.Router)();
router.post("/empresas", function (req, res) { return __awaiter(void 0, void 0, void 0, function () {
    var _a, empresas, relatorioEmpresa_1, formato, colunasSelecionadas_1, outputDir, fileName, filePath_1, workbook, sheet_1, pdfDoc_1, page_1, font_1, height_1, y_1, pdfBytes, error_1;
    return __generator(this, function (_b) {
        switch (_b.label) {
            case 0:
                _b.trys.push([0, 7, , 8]);
                _a = req.body, empresas = _a.empresas, relatorioEmpresa_1 = _a.relatorioEmpresa, formato = _a.formato;
                if (!empresas || !Array.isArray(empresas)) {
                    return [2 /*return*/, res.status(400).json({ error: "Empresas inválidas." })];
                }
                if (!relatorioEmpresa_1) {
                    return [2 /*return*/, res.status(400).json({ error: "Configuração de relatório ausente." })];
                }
                if (!["xlsx", "pdf"].includes(formato)) {
                    return [2 /*return*/, res.status(400).json({ error: "Formato inválido. Use 'xlsx' ou 'pdf'." })];
                }
                colunasSelecionadas_1 = Object.keys(relatorioEmpresa_1).filter(function (coluna) { return relatorioEmpresa_1[coluna] === true; });
                if (colunasSelecionadas_1.length === 0) {
                    return [2 /*return*/, res.status(400).json({ error: "Nenhuma coluna selecionada para o relatório." })];
                }
                outputDir = path_1.default.join(__dirname, "../../uploads");
                if (!fs_1.default.existsSync(outputDir))
                    fs_1.default.mkdirSync(outputDir);
                fileName = "relatorio_empresas_".concat(Date.now(), ".").concat(formato);
                filePath_1 = path_1.default.join(outputDir, fileName);
                if (!(formato === "xlsx")) return [3 /*break*/, 2];
                workbook = new exceljs_1.default.Workbook();
                sheet_1 = workbook.addWorksheet("Empresas");
                // Cabeçalhos
                sheet_1.addRow(colunasSelecionadas_1);
                // Dados
                empresas.forEach(function (empresa) {
                    var linha = colunasSelecionadas_1.map(function (col) { var _a; return (_a = empresa[col]) !== null && _a !== void 0 ? _a : ""; });
                    sheet_1.addRow(linha);
                });
                return [4 /*yield*/, workbook.xlsx.writeFile(filePath_1)];
            case 1:
                _b.sent();
                return [3 /*break*/, 6];
            case 2: return [4 /*yield*/, pdf_lib_1.PDFDocument.create()];
            case 3:
                pdfDoc_1 = _b.sent();
                page_1 = pdfDoc_1.addPage([595, 842]);
                return [4 /*yield*/, pdfDoc_1.embedFont(pdf_lib_1.StandardFonts.Helvetica)];
            case 4:
                font_1 = _b.sent();
                height_1 = page_1.getSize().height;
                y_1 = height_1 - 50;
                page_1.drawText("Relatório de Empresas", { x: 50, y: y_1, size: 16, font: font_1 });
                y_1 -= 30;
                empresas.forEach(function (empresa, i) {
                    colunasSelecionadas_1.forEach(function (col) {
                        var _a;
                        var texto = "".concat(col, ": ").concat((_a = empresa[col]) !== null && _a !== void 0 ? _a : "");
                        page_1.drawText(texto, { x: 50, y: y_1, size: 10, font: font_1 });
                        y_1 -= 15;
                    });
                    y_1 -= 10;
                    if (y_1 < 50) {
                        y_1 = height_1 - 50;
                        pdfDoc_1.addPage();
                    }
                });
                return [4 /*yield*/, pdfDoc_1.save()];
            case 5:
                pdfBytes = _b.sent();
                fs_1.default.writeFileSync(filePath_1, pdfBytes);
                _b.label = 6;
            case 6: return [2 /*return*/, res.download(filePath_1, fileName, function (err) {
                    if (err)
                        console.error("Erro ao enviar o arquivo:", err);
                    fs_1.default.unlinkSync(filePath_1); // remove após enviar
                })];
            case 7:
                error_1 = _b.sent();
                console.error(error_1);
                return [2 /*return*/, res.status(500).json({ error: "Erro ao gerar relatório." })];
            case 8: return [2 /*return*/];
        }
    });
}); });
// Rota de teste para confirmar que o backend está funcionando
router.get("/mensagem", function (req, res) {
    res.status(200).json({
        sucesso: true,
        mensagem: "✅ Backend de relatórios funcionando corretamente!",
    });
});
exports.default = router;
