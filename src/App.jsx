import ExcelChecker from './components/ExcelChecker'
import { CheckCircle2, XCircle, FileSpreadsheet, AlertTriangle, FileCheck, LayoutGrid } from 'lucide-react'

function App() {
  return (
    <div className="min-h-screen bg-white flex flex-col" style={{ fontFamily: 'Inter, system-ui, sans-serif' }}>
      {/* Top Header - Full Width */}
      <header className="bg-white border-b border-[#E5E7EB]">
        <div className="w-full mx-auto px-4 sm:px-5 lg:px-8 xl:px-12 py-2 sm:py-3 flex items-center justify-between max-w-[1600px]">
          <div className="flex items-center gap-2 sm:gap-3">
            <img
              src="/web_logo.png"
              alt="TPS Logo"
              className="h-8 sm:h-10 w-auto object-contain flex-shrink-0"
            />
            <div className="border-l border-[#E5E7EB] pl-2 sm:pl-3">
              <div className="text-[13px] sm:text-[15px] font-bold text-[#1A1A1A] tracking-wide uppercase leading-tight">HDC</div>
              <div className="text-[9px] sm:text-[10px] text-[#737373] leading-tight">Feature In The Ta'ang Population System</div>
            </div>
          </div>
          <div className="text-[9px] sm:text-[10px] text-[#737373] uppercase tracking-wider">v1.0</div>
        </div>
      </header>

      {/* Main Content - Full Width Responsive */}
      <main className="flex-1 w-full mx-auto px-3 sm:px-4 lg:px-6 xl:px-8 py-4 sm:py-6 lg:py-8 max-w-[1600px]">
        {/* Page Title Row - Full Width */}
        <div className="mb-6 sm:mb-8">
          <div className="border-b border-[#E5E7EB] pb-3 sm:pb-4">
            <div className="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
              <div>
                <p className="text-[9px] sm:text-[10px] text-[#737373] uppercase tracking-wider mb-1">အိမ်ထောင်စုများ database ထဲ မသွင်းရသေးခင် မှန်မှန်ကန်ကန်ရောက်ရှိရန်  စစ်ဆေး ခြင်းအတွက် အသုံးပြုသော Software ဖြစ်ပါသည်</p>
                <h1 className="text-[20px] sm:text-[24px] lg:text-[28px] font-semibold text-[#1A1A1A] tracking-tight">Excel Checker</h1>
                <p className="text-[11px] sm:text-[12px] text-[#737373] mt-1">Validate Excel/CSV files before database import</p>
              </div>
              <div className="hidden lg:block">
                <p className="text-[11px] text-[#737373]">Guidelines & Requirements</p>
              </div>
            </div>
          </div>
        </div>

        {/* Two Column Layout - Full Width */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 lg:gap-6 items-start">
          {/* Left Column - Excel Checker (wider) */}
          <div className="lg:col-span-7 xl:col-span-8">
            <ExcelChecker />
          </div>

          {/* Right Column - Guidelines Panel */}
          <aside className="lg:col-span-5 xl:col-span-4 space-y-3 sm:space-y-4">
            {/* File Requirements */}
            <section className="border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
              <div className="border-b border-[#E5E7EB] p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
                    <FileCheck size={14} className="sm:w-4 sm:h-4 text-[#1A1A1A]" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-[#1A1A1A]">File Requirements</h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <ul className="space-y-1 text-[10px] sm:text-[11px] text-[#737373]">
                  <li>• Supported: .XLSX, .XLS, .CSV</li>
                  <li>• Max size: 10MB</li>
                  <li>• First row must contain headers</li>
                </ul>
              </div>
            </section>

            {/* DO Section */}
            <section className="border border-green-200 bg-green-50/20" style={{ borderRadius: '0px' }}>
              <div className="border-b border-green-200 p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-green-100 flex items-center justify-center flex-shrink-0">
                    <CheckCircle2 size={14} className="sm:w-4 sm:h-4 text-green-600" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-green-800">Do / လုပ်ရန်</h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <ul className="space-y-1 text-[10px] sm:text-[11px] text-[#1A1A1A]">
                  <li className="flex items-start gap-2">
                    <span className="text-green-600 font-bold">✓</span>
                    <span>Use Unicode Myanmar font</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-green-600 font-bold">✓</span>
                    <span>Fill all required fields</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-green-600 font-bold">✓</span>
                    <span>Check spelling before upload</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-green-600 font-bold">✓</span>
                    <span>Use DD-MM-YYYY date format</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-green-600 font-bold">✓</span>
                    <span>Verify household numbers</span>
                  </li>
                </ul>
              </div>
            </section>

            {/* DON'T Section */}
            <section className="border border-red-200 bg-red-50/20" style={{ borderRadius: '0px' }}>
              <div className="border-b border-red-200 p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-red-100 flex items-center justify-center flex-shrink-0">
                    <XCircle size={14} className="sm:w-4 sm:h-4 text-red-600" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-red-800">Don't / ရှောင်ရန်</h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <ul className="space-y-1 text-[10px] sm:text-[11px] text-[#1A1A1A]">
                  <li className="flex items-start gap-2">
                    <span className="text-red-600 font-bold">✗</span>
                    <span>Leave required fields empty</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-red-600 font-bold">✗</span>
                    <span>Mix Zawgyi & Unicode fonts</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-red-600 font-bold">✗</span>
                    <span>Use special characters</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-red-600 font-bold">✗</span>
                    <span>Merge cells in Excel</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-red-600 font-bold">✗</span>
                    <span>Add duplicate entries</span>
                  </li>
                </ul>
              </div>
            </section>

            {/* Required Fields */}
            <section className="border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
              <div className="border-b border-[#E5E7EB] p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
                    <LayoutGrid size={14} className="sm:w-4 sm:h-4 text-[#1A1A1A]" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-[#1A1A1A]">Required Fields / မဖြစ်မနေဖြည့်သွင်းရန် လိုအပ်သည် </h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <div className="grid grid-cols-2 gap-x-2 sm:gap-x-3 gap-y-1 text-[10px] sm:text-[11px]">
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">Ward/Village/Group</span>
                  </div>
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">Township</span>
                  </div>
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">District</span>
                  </div>
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">Gender</span>
                  </div>
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">Relationship</span>
                  </div>
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="w-1.5 h-1.5 bg-orange-500 flex-shrink-0"></span>
                    <span className="text-[#737373]">Name (Myanmar)</span>
                  </div>
                </div>
              </div>
            </section>
          </aside>
        </div>
      </main>

      {/* Footer - Full Width */}
      <footer className="border-t border-[#E5E7EB] mt-auto">
        <div className="w-full mx-auto px-3 sm:px-4 lg:px-6 xl:px-8 max-w-[1600px]">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 py-3 sm:py-4">
            <div>
              <p className="text-[10px] sm:text-[11px] text-[#737373]">Household Database Checker — Data validation tool</p>
              <p className="text-[10px] sm:text-[11px] text-[#737373]">© 2026 Ta'ang Population System</p>
            </div>
            <div className="text-right">
              <p className="text-[9px] sm:text-[10px] text-[#ABABAB] uppercase tracking-wider">Designed & Developed by</p>
              <p className="text-[10px] sm:text-[11px] text-[#737373] font-medium">Mai San Hlu & Mai Nay Lin</p>
            </div>
          </div>
        </div>
      </footer>
    </div>
  )
}

export default App
