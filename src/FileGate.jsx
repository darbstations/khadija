import { useState } from "react";
import { setFiles, installFsShim } from "./fileStore";

const FONT = "'Segoe UI', Tahoma, 'Arial', sans-serif";

// Gate shown before the dashboard: collects the .xlsx files, populates the
// in-memory store, installs the window.fs shim, then hands control to App.
export default function FileGate({ onReady }) {
  const [picked, setPicked] = useState([]);
  const [dragging, setDragging] = useState(false);

  function accept(fileList) {
    const xlsx = Array.from(fileList).filter(f => /\.xlsx?$/i.test(f.name));
    if (xlsx.length) setPicked(xlsx);
  }

  function start() {
    setFiles(picked);
    installFsShim();
    onReady();
  }

  return (
    <div dir="rtl" style={{ fontFamily: FONT }} className="min-h-screen bg-slate-950 text-white flex items-center justify-center p-4">
      <div className="w-full max-w-xl">
        <h1 className="text-2xl font-bold mb-1">⛽ داش بورد محطات درب الوقود</h1>
        <p className="text-slate-400 text-sm mb-6">ارفعي ملفات Excel (.xlsx) الخاصة بالمحطات لبدء التحليل — مارس / أبريل / مايو 2026</p>

        <label
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={e => { e.preventDefault(); setDragging(false); accept(e.dataTransfer.files); }}
          className={`block cursor-pointer rounded-xl border-2 border-dashed p-10 text-center transition ${dragging ? "border-sky-400 bg-sky-500/10" : "border-slate-700 bg-slate-900 hover:border-slate-500"}`}
        >
          <div className="text-4xl mb-3">📁</div>
          <div className="font-bold mb-1">اسحبي الملفات هنا أو اضغطي للاختيار</div>
          <div className="text-slate-400 text-sm">يمكن اختيار عدة ملفات .xlsx دفعة واحدة</div>
          <input
            type="file"
            accept=".xlsx,.xls"
            multiple
            className="hidden"
            onChange={e => accept(e.target.files)}
          />
        </label>

        {picked.length > 0 && (
          <div className="mt-4 bg-slate-900 border border-slate-800 rounded-xl p-4">
            <div className="text-emerald-400 text-sm font-bold mb-2">✅ تم اختيار {picked.length} ملف:</div>
            <div className="max-h-40 overflow-y-auto text-xs text-slate-400 space-y-1">
              {picked.map((f, i) => <div key={i}>• {f.name}</div>)}
            </div>
          </div>
        )}

        <button
          onClick={start}
          disabled={picked.length === 0}
          className="mt-5 w-full px-4 py-3 rounded-lg text-sm font-bold transition bg-sky-500 hover:bg-sky-600 disabled:bg-slate-800 disabled:text-slate-500 disabled:cursor-not-allowed text-white"
        >
          🚀 بدء التحليل
        </button>

        <p className="text-slate-600 text-xs mt-4 leading-relaxed">
          ملاحظة: تتم معالجة الملفات داخل المتصفح فقط ولا يتم رفعها إلى أي خادم. أسماء الملفات يجب أن تطابق الأكواد المعرّفة في الداش بورد (مثل
          <span className="text-slate-400" dir="ltr"> "MK072 3.xlsx"</span>).
        </p>
      </div>
    </div>
  );
}
