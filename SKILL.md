---
name: ppt-map-generator
description: Generate PowerPoint presentations with map boundaries for Chinese administrative divisions
---

## Use This Skill

Put this line at the top of your prompt:

```
@ppt-map-generator
```

Then specify which city/district you want to generate maps for.

## Overview

This skill generates PowerPoint presentations containing map boundary decompositions for Chinese administrative divisions (provinces, cities, districts, counties).

## Key Technical Insight

**python-pptx cannot insert vector graphics** - It only supports bitmap formats (PNG, JPEG).

**Solution**: Use Windows COM automation (`win32com.client`) to insert SVG files directly into PowerPoint. PowerPoint 2019/365 natively supports SVG as true vector graphics.

## Data Source

- Primary: Alibaba DataV GeoAtlas API
- URL format: `https://geo.datav.aliyun.com/areas_v3/bound/{adcode}_full.json`
- Example: China (全国) - adcode 100000
- Example: Zhejiang Province (浙江省) - adcode 330000

## Required Libraries

```bash
pip install matplotlib requests pywin32 numpy
```

**Important**: `pywin32` is required for Windows COM automation to insert SVG vectors.

## Output Files

```
output/
├── china_provinces_map.pptx    # Main presentation (with vector SVG)
├── maps/
│   ├── china_overview.svg      # Overview map
│   ├── 北京市.svg              # Individual province maps
│   ├── 上海市.svg
│   └── ...
```

## Implementation Pattern

```python
import win32com.client
import matplotlib.pyplot as plt

# 1. Create SVG with matplotlib
fig, ax = plt.subplots(figsize=(10, 8))
# ... draw map geometry ...
fig.savefig('map.svg', format='svg')

# 2. Insert SVG via COM automation
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt_app.Presentations.Add()
slide = presentation.Slides.AddSlide(1, blank_layout)
shape = slide.Shapes.AddPicture(svg_path, False, True, left, top, width, height)

# 3. Save
presentation.SaveAs(output_path)
```

## Slide Layout

1. **Cover Slide** - Title and subtitle
2. **Overview Slide** - All regions (without legend, to keep clean vector)
3. **Region Slides** - Individual region maps (one per slide)

## Important Implementation Notes

### SVG Content

- **Do NOT include titles in SVG** - Titles should be separate text boxes in PowerPoint
- **Do NOT include legends in SVG** - Legends extend beyond map bounds and break layout
- Maps should contain only the geometric shapes (polygons)

### SVG Figure Size

Use **large figure dimensions** (20-28 inches) for high-quality vector output:

```python
# Individual province maps
fig_width = 20  # Large size for quality
fig_height = fig_width * (height / width)

# China overview map
fig_width = 28  # Even larger for multi-province map
```

This ensures SVG files have sufficient detail when scaled in PowerPoint.

### Output File Naming

Use **sequence numbers** to preserve history:

```python
def get_next_pptx_path(output_dir, base_name="china_provinces_map"):
    existing = [f for f in os.listdir(output_dir) 
                if f.startswith(base_name) and f.endswith(".pptx")]
    # ... extract max sequence number ...
    next_seq = max(seq_nums) + 1
    return os.path.join(output_dir, f"{base_name}_{next_seq:02d}.pptx")
```

Output: `china_provinces_map_01.pptx`, `china_provinces_map_02.pptx`, etc.

### Slide Sizing

**CRITICAL**: Always preserve aspect ratio when inserting images.

**Rule**: Insert images at **80% of slide size**, centered on slide.

```python
# 1. Get slide dimensions (in points)
slide_width = presentation.PageSetup.SlideWidth   # ~960 points (13.33")
slide_height = presentation.PageSetup.SlideHeight # ~540 points (7.5")

# 2. Calculate 80% of slide size
available_width = int(slide_width * 0.8)   # ~768 points
available_height = int(slide_height * 0.8) # ~432 points

# 3. Insert at original size first (no width/height params)
shape = slide.Shapes.AddPicture(svg_path, False, True, 0, 0)

# 4. Get original dimensions
original_width = shape.Width
original_height = shape.Height
original_ratio = original_width / original_height

# 5. Scale proportionally to fit 80% of slide
available_ratio = available_width / available_height

if original_ratio > available_ratio:
    # Image is wider - fit by width
    new_width = available_width
    new_height = available_width / original_ratio
else:
    # Image is taller - fit by height
    new_height = available_height
    new_width = available_height * original_ratio

# 6. Apply new size and center on slide
shape.Width = int(new_width)
shape.Height = int(new_height)
shape.Left = int((slide_width - new_width) / 2)
shape.Top = int((slide_height - new_height) / 2)
```

**WRONG** - This distorts the image:
```python
# DO NOT DO THIS - breaks aspect ratio!
shape = slide.Shapes.AddPicture(path, False, True, left, top, width, height)
```

## Color Palette

```python
COLOR_PALETTE = [
    "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7",
    "#DDA0DD", "#98D8C8", "#F7DC6F", "#BB8FCE", "#FF9FF3",
    "#54A0FF", "#5F27CD", "#00D2D3", "#FF6B81", "#7BED9F",
    "#A29BFE", "#FD79A8", "#FDCB6E", "#6C5CE7", "#E17055",
    "#00B894", "#0984E3", "#6C5B7B", "#F8B500", "#B8E994",
]
```

## Working with Vector Graphics in PowerPoint

### After PPTX is Generated

1. Select the map image
2. Right-click → **"Convert to Shape"** (转换为形状)
3. Or Right-click → **"Ungroup"** (取消组合)
4. Each region becomes an editable shape

### Manual SVG Import

If you need to manually import SVG files:

1. In PowerPoint: Insert → Pictures → This Device
2. Select the SVG file
3. Right-click the image → Convert to Shape

## Common Administrative Divisions

### National Level
- China (全国): 100000

### Province Level Examples
- Beijing (北京市): 110000
- Shanghai (上海市): 310000
- Guangdong (广东省): 440000
- Zhejiang (浙江省): 330000
- Sichuan (四川省): 510000
- Xinjiang (新疆维吾尔自治区): 650000
- Tibet (西藏自治区): 540000
- Taiwan (台湾省): 710000
- Hong Kong (香港特别行政区): 810000
- Macau (澳门特别行政区): 820000

### Finding Adcodes

1. Visit: `https://geo.datav.aliyun.com/areas_v3/bound/{adcode}_full.json`
2. The `_full` suffix returns all child divisions
3. Without `_full` returns just that division's boundary

## Usage Examples

### Generate China provinces map

```
@ppt-map-generator
Generate a PowerPoint presentation with map boundaries for all provinces in China.
```

### Generate city districts map

```
@ppt-map-generator
Generate maps for Hangzhou City (杭州市, adcode: 330100) with all districts.
```

### Custom colors

```
@ppt-map-generator
Generate maps for Zhejiang Province with custom colors:
- Hangzhou: #FF6B6B
- Ningbo: #4ECDC4
- Wenzhou: #45B7D1
```

## Troubleshooting

### Issue: Images not appearing in PPTX

**Cause**: COM automation may fail silently

**Solution**: Ensure PowerPoint is installed and accessible via COM. Check that SVG files exist and paths are absolute.

### Issue: Chinese characters not displaying

**Solution**: Configure matplotlib fonts
```python
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
```

### Issue: SVG appears but cannot ungroup

**Cause**: PowerPoint version doesn't support SVG

**Solution**: Requires PowerPoint 2019 or Microsoft 365. Older versions need EMF format conversion via Inkscape.

### Issue: Unicode encoding errors on Windows

**Solution**: Configure stdout encoding
```python
import sys
sys.stdout.reconfigure(encoding='utf-8')
```

### Issue: Image extends beyond slide canvas

**Cause**: Image position/size calculated incorrectly

**Solution**: Calculate available space accounting for title and margins. See "Slide Sizing" section above.

### Issue: Title/legend appears in map image

**Cause**: matplotlib `ax.set_title()` or `ax.legend()` rendered in SVG

**Solution**: Remove title/legend from matplotlib figure. Add title as separate text box in PowerPoint.

### Issue: Image aspect ratio distorted

**Cause**: Specifying both width AND height in `AddPicture()` stretches image

**Solution**: Insert at original size first, then scale proportionally. See "Slide Sizing" section above.

## Best Practices

1. **Always use absolute paths** for COM automation
2. **Cache GeoJSON data** to avoid repeated API calls
3. **Use consistent color palette** across related maps
4. **Set proper aspect ratio** based on map bounds
5. **Insert at 80% of slide size** for optimal display
6. **Use sequence numbers** in output filenames for version tracking

## Debugging Tips

### Print slide and image dimensions

```python
print(f"Slide size: {slide_width} x {slide_height}")
print(f"80% size: {int(slide_width * 0.8)} x {int(slide_height * 0.8)}")
print(f"Inserted: {svg_path} (size: {new_width}x{new_height})")
```

This helps verify that images are being sized correctly.

## Template Script

A complete working script is available at:
`output/generate_china_maps.py`

This script demonstrates:
- Fetching GeoJSON from DataV API
- Creating SVG maps with matplotlib
- Inserting SVG via Windows COM
- Proper aspect ratio handling
- Color palette generation

## Related Skills

- @data-visualization - For charts and graphs
- @python-automation - For scripting repetitive tasks

## Key Learnings from Development

### Why python-pptx Doesn't Work for Vectors

The `python-pptx` library uses PIL to process images, which only supports bitmap formats. Even if you try to insert SVG or EMF, it gets converted to a bitmap, losing vector editability.

### Why EMF Creation via GDI Failed

Creating EMF files using Windows GDI (`CreateEnhMetaFileW`) produces valid files, but the header information was incomplete, making the files unusable in PowerPoint.

### Why SVG via COM Works

PowerPoint 2019/365 has native SVG support. Using COM automation bypasses python-pptx's PIL conversion and inserts SVG directly as a vector object that can be ungrouped and edited.

### Common Mistakes to Avoid

1. **Don't specify width+height together** - Breaks aspect ratio
2. **Don't put titles/legends in SVG** - Affects layout and sizing
3. **Don't use small figure sizes** - Results in low-quality output
4. **Don't overwrite output files** - Use sequence numbers for comparison