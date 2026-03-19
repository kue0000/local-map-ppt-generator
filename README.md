# Map PPT Generator - OpenCode Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)

Generate PowerPoint presentations with vector map boundaries for Chinese administrative divisions (provinces, cities, districts, counties).

## Features

- 🗺️ Generate high-quality vector maps (SVG) for any Chinese administrative division
- 📊 Create PowerPoint presentations with editable vector graphics
- 🎨 Support for custom colors and color palettes
- 📱 Works with PowerPoint 2019/365 (native SVG support)
- 🔄 Right-click to ungroup and edit individual map regions

## Quick Start

### Prerequisites

- Python 3.7+
- Microsoft PowerPoint 2019 or later (for SVG support)
- Windows OS (required for COM automation)

### Installation

```bash
# Clone this repository
git clone https://github.com/kue0000/china-map-ppt-generator.git
cd china-map-ppt-generator

# Install dependencies
pip install -r requirements.txt
```

### Usage

#### Generate China Provinces Map

```bash
python scripts/generate_china_maps.py
```

This will create:
- `china_provinces_map_01.pptx` - PowerPoint presentation with all provinces
- `maps/` - Directory containing SVG files for each province

#### Generate Hangzhou Districts Map

```bash
python scripts/generate_hangzhou_maps.py
```

This will create:
- `hangzhou_districts_map_01.pptx` - PowerPoint presentation with all districts
- `hangzhou_maps/` - Directory containing SVG files for each district

#### Generate Custom City Maps

Modify the `adcode` in the script to generate maps for any city:

```python
# Example: Generate maps for Shanghai (adcode: 310000)
SHANGHAI_ADCODE = "310000"
```

Common adcodes:
- Beijing (北京市): 110000
- Shanghai (上海市): 310000
- Guangdong (广东省): 440000
- Zhejiang (浙江省): 330000
- Hangzhou (杭州市): 330100

Find more adcodes at: [DataV GeoAtlas](https://datav.aliyun.com/portal/school/atlas/area_selector)

## How It Works

1. **Fetch GeoJSON**: Downloads geographic boundary data from Alibaba DataV GeoAtlas API
2. **Generate SVG**: Creates SVG vector maps using matplotlib
3. **Create PPTX**: Inserts SVG files into PowerPoint using Windows COM automation
4. **Vector Editing**: Maps remain as editable vectors in PowerPoint (right-click → Ungroup)

## Output Structure

```
├── china_provinces_map_01.pptx    # Main presentation
├── maps/
│   ├── china_overview.svg         # Overview map
│   ├── 北京市.svg                 # Individual province maps
│   ├── 上海市.svg
│   └── ...
├── scripts/
│   ├── generate_china_maps.py     # Main generation script
│   └── generate_hangzhou_maps.py  # Hangzhou example
├── SKILL.md                       # OpenCode skill documentation
└── README.md                      # This file
```

## Working with Vector Graphics in PowerPoint

After generating the PPTX file:

1. Select any map image
2. Right-click → **"Convert to Shape"** (转换为形状)
3. Or Right-click → **"Ungroup"** (取消组合)
4. Each region becomes an editable shape

This allows you to:
- Change colors of individual regions
- Add labels and annotations
- Modify boundaries
- Export as other formats

## Customization

### Custom Colors

Modify the `COLOR_PALETTE` in the script:

```python
COLOR_PALETTE = [
    "#FF6B6B",  # Red
    "#4ECDC4",  # Teal
    "#45B7D1",  # Blue
    # ... add more colors
]
```

### Map Style

Adjust drawing parameters:

```python
# Line width
linewidth = 1.0

# Transparency
alpha = 0.7

# Edge color
edge_color = "black"
```

## Data Source

This project uses the [Alibaba DataV GeoAtlas API](https://datav.aliyun.com/portal/school/atlas/area_selector) for geographic boundary data.

API URL format: `https://geo.datav.aliyun.com/areas_v3/bound/{adcode}_full.json`

## Requirements

- **Python**: 3.7 or higher
- **PowerPoint**: 2019 or Microsoft 365 (for SVG support)
- **OS**: Windows (required for COM automation)

### Python Packages

- `matplotlib` - Map generation and SVG creation
- `requests` - GeoJSON data fetching
- `pywin32` - Windows COM automation
- `numpy` - Coordinate processing

## Troubleshooting

### Issue: Images not appearing in PPTX

**Cause**: COM automation may fail silently

**Solution**: 
- Ensure PowerPoint is installed and accessible
- Check that SVG files exist and paths are absolute
- Run script as administrator if needed

### Issue: Chinese characters not displaying

**Solution**: Configure matplotlib fonts:

```python
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
```

### Issue: SVG appears but cannot ungroup

**Cause**: PowerPoint version doesn't support SVG

**Solution**: Requires PowerPoint 2019 or Microsoft 365. Older versions need EMF format conversion.

### Issue: Unicode encoding errors on Windows

**Solution**: Configure stdout encoding:

```python
import sys
sys.stdout.reconfigure(encoding='utf-8')
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Alibaba DataV GeoAtlas](https://datav.aliyun.com/portal/school/atlas/area_selector) for geographic data
- [matplotlib](https://matplotlib.org/) for map generation
- [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint automation

## Author

**kue0000**
- GitHub: [@kue0000](https://github.com/kue0000)
- Email: 1576092480@qq.com

## Star History

If you find this project useful, please consider giving it a star! ⭐
