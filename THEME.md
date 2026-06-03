# HDC - Design Theme Documentation

## Overview

HDC (Household Database Checker) uses a **minimal, professional design system** built with Tailwind CSS. The theme emphasizes clarity, accessibility, and bilingual support (English + Myanmar).

---

## Design Philosophy

### Core Principles
1. **Minimal & Clean** - No unnecessary decorations
2. **Professional** - Government/enterprise aesthetic
3. **Accessible** - High contrast, readable typography
4. **Responsive** - Works on mobile, tablet, desktop
5. **Sharp Corners** - `border-radius: 0px` throughout (deliberate choice)

---

## Color Palette

### Primary Colors

| Color | Hex Code | Usage |
|-------|----------|-------|
| **Black** | `#1A1A1A` | Primary text, headings |
| **White** | `#FFFFFF` | Background, cards |
| **Gray 100** | `#F3F4F6` | Secondary backgrounds |
| **Gray 200** | `#E5E7EB` | Borders, dividers |
| **Gray 500** | `#737373` | Secondary text, placeholders |
| **Gray 400** | `#ABABAB` | Muted text, footer |

### Accent Colors

#### Success (Green)
| Color | Hex Code | Usage |
|-------|----------|-------|
| **Green 50** | `#F0FDF4` | Success banners background |
| **Green 100** | `#BBF7D0` | Light green accents |
| **Green 200** | `#86EFAC` | Borders, highlights |
| **Green 400** | `#4ADE80` | Checkmarks, icons |
| **Green 600** | `#16A34A` | Primary success, borders |
| **Green 800** | `#166534` | Success headings |
| **Green 900** | `#14532D` | Dark success text |

#### Error (Red)
| Color | Hex Code | Usage |
|-------|----------|-------|
| **Red 50** | `#FEF2F2` / `red-50/20` | Error backgrounds |
| **Red 100** | `#FEE2E2` | Error icon backgrounds |
| **Red 200** | `#FECACA` | Error borders |
| **Red 600** | `#DC2626` | Error text, icons, X marks |
| **Red 800** | `#991B1B` | Error headings |

#### Warning (Orange/Amber)
| Color | Hex Code | Usage |
|-------|----------|-------|
| **Orange 500** | `#F97316` | Optional field indicators |

---

## Typography

### Font Family
```css
font-family: 'Inter', system-ui, sans-serif;
```

### Font Sizes

| Element | Size (Mobile) | Size (Desktop) | Weight | Color |
|---------|---------------|----------------|--------|-------|
| **Page Title** | 20px | 28px | 600 (semibold) | `#1A1A1A` |
| **Section Title** | 12px | 13px | 600 (semibold) | `#1A1A1A` |
| **Card Title** | 14px | 14px | 600 (semibold) | `#1A1A1A` |
| **Body Text** | 10px | 11px | 400 (normal) | `#737373` |
| **Label/Meta** | 9px | 10px | 500 (medium) | `#737373` |
| **Stats/Numbers** | 11px | 12px | 500 (medium) | Various |
| **Button Text** | 11px | 12px | 500 (medium) | `#1A1A1A` |

### Myanmar Text Support
- Unicode Myanmar fonts (Padauk, Myanmar3)
- Same sizing as English text
- Proper line-height for Myanmar scripts

---

## Spacing System

### Padding Scale
| Token | Value | Usage |
|-------|-------|-------|
| **p-2** | 8px | Compact card padding |
| **p-2.5** | 10px | Standard card padding |
| **p-3** | 12px | Section padding |
| **p-4** | 16px | Content area padding |
| **p-4 sm:p-6** | 16px/24px | Page padding |

### Gap Scale
| Token | Value | Usage |
|-------|-------|-------|
| **gap-1** | 4px | Tight spacing |
| **gap-1.5** | 6px | List items |
| **gap-2** | 8px | Standard spacing |
| **gap-2.5** | 10px | Card headers |
| **gap-3** | 12px | Section spacing |
| **gap-4** | 16px | Major sections |

### Container
```css
max-width: 1600px;
padding: 12px / 16px / 24px / 32px (responsive);
```

---

## Borders & Shadows

### Border Radius
**NO rounded corners** - sharp corners throughout:
```jsx
style={{ borderRadius: '0px' }}
// or Tailwind: rounded-none
```

This is a deliberate design choice for a professional, government-style aesthetic.

### Border Colors
| Usage | Color | Width |
|-------|-------|-------|
| **Default** | `#E5E7EB` (gray-200) | 1px |
| **Success** | `#16A34A` (green-600) | 1px |
| **Error** | `#FECACA` (red-200) | 1px |
| **Dashed (Upload)** | `#E5E7EB` | 2px |

### Shadows
**NO shadows** - flat design throughout:
- No box-shadow on cards
- No elevation effects
- Pure flat aesthetic

---

## Components

### Cards / Sections
```jsx
<section className="border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
  <div className="border-b border-[#E5E7EB] p-2.5 sm:p-3">
    {/* Header with icon */}
  </div>
  <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
    {/* Content */}
  </div>
</section>
```

**Card Variants:**
- **Default**: White background, gray border
- **Success**: `bg-green-50/20`, `border-green-200`
- **Error**: `bg-red-50/20`, `border-red-200`

### Icons
- **Library**: Lucide React
- **Sizes**: 14px (default), 16px (large), 18px (xl)
- **Icon Containers**: 28x28px or 32x32px squares with background

```jsx
<div className="w-7 h-7 sm:w-8 sm:h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
  <Icon size={14} className="sm:w-4 sm:h-4 text-[#1A1A1A]" />
</div>
```

### Buttons
```jsx
<button className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 
  bg-white text-[#1A1A1A] border border-[#E5E7EB] 
  hover:bg-[#F3F4F6] transition-colors 
  text-[11px] sm:text-[12px] font-medium"
  style={{ borderRadius: '0px' }}>
  <Icon size={14} />
  Button Text
</button>
```

**Button Variants:**
- Primary: White background, gray border
- Hover: Gray-100 background

### Lists (DO's & DON'Ts)
```jsx
<ul className="space-y-1.5 text-[10px] sm:text-[11px]">
  <li className="flex items-start gap-1.5">
    <span className="text-green-600 font-bold mt-0.5">✓</span>
    <span>
      <span className="font-medium">Myanmar Text</span>
      <span className="text-[#737373]"> (English)</span>
    </span>
  </li>
</ul>
```

### Upload Area
```jsx
<label className="flex flex-col items-center justify-center w-full h-36 sm:h-44 
  border border-dashed border-[#E5E7EB] cursor-pointer bg-white 
  hover:bg-[#F3F4F6] transition-colors px-4" 
  style={{ borderWidth: '2px' }}>
  {/* Upload icon and text */}
</label>
```

---

## Layout Grid

### Page Structure
```
┌─────────────────────────────────────────┐
│ Header (Logo + Title + Version)         │
├─────────────────────────────────────────┤
│ Page Title Row                          │
├─────────────────────────────────────────┤
│                                         │
│  Excel Checker Component (Full Width)   │
│                                         │
│  ┌─────────────────────────────────┐    │
│  │ Upload Area                     │    │
│  └─────────────────────────────────┘    │
│                                         │
│  ┌─────────┬─────────┬─────────────┐    │
│  │ File    │   DO    │   Required  │    │
│  │ Req     │         │   Fields    │    │
│  └─────────┴─────────┴─────────────┘    │
│                                         │
│  ┌─────────────────────────────────┐    │
│  │ DON'T Section (Full Width)      │    │
│  └─────────────────────────────────┘    │
│                                         │
├─────────────────────────────────────────┤
│ Footer (Credits)                        │
└─────────────────────────────────────────┘
```

### Responsive Breakpoints
| Breakpoint | Width | Layout Changes |
|------------|-------|----------------|
| **Mobile** | < 640px | Single column, stacked layout |
| **Tablet** | 640px - 1024px | 2-column guidelines grid |
| **Desktop** | > 1024px | Full layout, 3-column guidelines |

---

## Animations

### Success Animation (Lamp-inspired)
When all checks pass, a subtle green light beam animation plays:

```jsx
<motion.div
  initial={{ opacity: 0, y: 6 }}
  animate={{ opacity: 1, y: 0 }}
  transition={{ duration: 0.45, ease: 'easeOut' }}
  className="relative overflow-hidden border border-[#16A34A] bg-[#F0FDF4]"
>
  {/* Conic gradient lamp glow */}
  <motion.div
    initial={{ opacity: 0, scaleX: 0.3 }}
    animate={{ opacity: [0, 0.55, 0.18] }}
    transition={{ duration: 1.1, ease: 'easeOut', delay: 0.1 }}
    style={{
      background: 'conic-gradient(from 250deg at 50% 0%, transparent 0deg, #4ADE80 30deg, #86EFAC 60deg, transparent 120deg)',
      filter: 'blur(18px)',
    }}
  />
  
  {/* Horizontal scan line sweep */}
  <motion.div
    initial={{ x: '-100%', opacity: 0.9 }}
    animate={{ x: '120%', opacity: 0 }}
    transition={{ duration: 0.85, ease: 'easeInOut', delay: 0.2 }}
  />
</motion.div>
```

**Animation Sequence (~1.1s):**
1. Card fades up (0.45s)
2. Conic lamp glow bloom (1.1s)
3. Scan line sweep left→right (0.85s)
4. Check icon springs in (0.4s, bounce)
5. Text slides in (0.38s)
6. Stats row fades in last (0.35s)

---

## Bilingual Design

### Text Pattern
All UI text follows bilingual pattern:
```
[Myanmar Text] ([English Text])
```

Examples:
- `"လုပ်ရန်" (Do)`
- `"ရှောင်ရန်" (Don't)`
- `"မြို့နယ်" (Township)`

### Font Weight Hierarchy
- **Myanmar text**: `font-medium` (500)
- **English translation**: `text-[#737373]` (gray-500), normal weight

### Layout Considerations
- Myanmar text may wrap differently than English
- Allow extra vertical space for stacked text
- Test with longest Myanmar words

---

## Accessibility

### Contrast Ratios
- Primary text (#1A1A1A) on white: **12.6:1** (AAA)
- Secondary text (#737373) on white: **4.6:1** (AA)
- Green on white: **3.1:1** (AA Large)

### Focus States
- Visible focus indicators on interactive elements
- Tab navigation support

### Screen Reader Support
- Semantic HTML structure
- ARIA labels where needed
- Icon-only buttons have aria-labels

---

## File Structure

```
src/
├── components/
│   └── ExcelChecker.jsx    # Main component with theme
├── App.jsx                 # Layout wrapper
└── index.css              # Global styles (if any)
```

### Tailwind Config
```javascript
// tailwind.config.js
module.exports = {
  theme: {
    extend: {
      colors: {
        // Custom colors if needed
      },
      keyframes: {
        fadeInUp: {
          '0%': { opacity: '0', transform: 'translateY(8px)' },
          '100%': { opacity: '1', transform: 'translateY(0)' },
        },
      },
      animation: {
        'fade-in-up': 'fadeInUp 0.35s ease-out both',
      },
    },
  },
}
```

---

## Usage Guidelines

### Do's
✅ Use sharp corners (`rounded-none` or `borderRadius: '0px'`)  
✅ Maintain high contrast for readability  
✅ Use consistent spacing (multiples of 4px)  
✅ Include Myanmar translations for all text  
✅ Test responsive behavior at all breakpoints  
✅ Use the color palette consistently

### Don'ts
❌ Add rounded corners or shadows  
❌ Use colors outside the defined palette  
❌ Skip Myanmar translations  
❌ Use different font sizes than defined  
❌ Add decorative elements  

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| v1.0 | June 2026 | Initial theme documentation |

---

**Design System by**: Mai Naung Naung & Mai Nay Lin  
**Project**: HDC - Household Database Checker
