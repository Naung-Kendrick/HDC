import { motion } from 'framer-motion'
import ExcelChecker from './components/ExcelChecker'

function App() {
  return (
    <div className="min-h-screen bg-white flex flex-col" style={{ fontFamily: 'Inter, system-ui, sans-serif' }}>
      {/* Top Header - Full Width */}
      <header className="bg-white border-b border-[#E5E7EB]">
        <div className="w-full mx-auto px-4 sm:px-5 lg:px-8 xl:px-12 py-2 sm:py-3 flex items-center justify-between max-w-[1600px]">
          <div className="flex items-center gap-3 sm:gap-4">
            <img
              src="/web_logo.png"
              alt="TLFUG Immigration Logo"
              className="h-9 sm:h-11 w-9 sm:w-11 object-contain flex-shrink-0"
            />
            <div className="border-l border-[#E5E7EB] pl-3 sm:pl-4">
              <div className="text-[13px] sm:text-[15px] font-bold text-[#1A1A1A] tracking-wide uppercase leading-tight">HDC</div>
              <div className="text-[9px] sm:text-[10px] text-[#737373] leading-tight">Household Database Checker  </div>
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
                <motion.p 
                  className="text-[10px] sm:text-[11px] text-red-600 font-bold uppercase tracking-wider mb-1 bg-red-50 border border-red-200 px-3 py-2 w-full block"
                  animate={{ 
                    boxShadow: ['0 0 0 0 rgba(220, 38, 38, 0)', '0 0 0 6px rgba(220, 38, 38, 0.2)', '0 0 0 0 rgba(220, 38, 38, 0)']
                  }}
                  transition={{ 
                    duration: 2,
                    repeat: Infinity,
                    repeatType: 'loop'
                  }}
                >အိမ်ထောင်စုများ database ထဲ မသွင်းရသေးမှီ သတ်မှတ်ထားသော စံနှုန်းများ (Excel standard frame )အတိုင်း မှန်မှန်ကန်ကန်ရောက်ရှိစေရန်  စစ်ဆေး ခြင်းအတွက် အသုံးပြုသော Software ဖြစ်ပါသည်...အသုံးမပြုမှီ လိုက်နာရမည့်အချက်များကို သေချာရှင်းလင်းစွာ ဖတ်ရှုလိုက်နာပါ။</motion.p>
                <h1 className="text-[20px] sm:text-[24px] lg:text-[28px] font-semibold text-[#1A1A1A] tracking-tight">Excel Checker</h1>
                <p className="text-[11px] sm:text-[12px] text-[#737373] mt-1">Validate Excel/CSV files before database import</p>
              </div>
            </div>
          </div>
        </div>

        {/* Excel Checker - Full Width */}
        <ExcelChecker />
      </main>

      {/* Footer - Full Width */}
      <footer className="border-t border-[#E5E7EB] mt-auto">
        <div className="w-full mx-auto px-3 sm:px-4 lg:px-6 xl:px-8 max-w-[1600px]">
          <div className="flex justify-end items-center py-3 sm:py-4">
            <div className="text-right">
              <p className="text-[8px] sm:text-[9px] text-[#ABABAB] uppercase tracking-widest font-medium">Designed & Developed by</p>
              <p className="text-[11px] sm:text-[12px] text-[#1A1A1A] font-semibold tracking-tight">Mai Naung Naung <span className="text-[#ABABAB] font-normal">&</span> Mai Nay Lin</p>
            </div>
          </div>
        </div>
      </footer>
    </div>
  )
}

export default App
