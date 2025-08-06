#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最終版 PDF 抽取器 - 支援所有格式的工單明細表
整合固定格式解析和增強格式匹配
"""

import pdfplumber
import pandas as pd
import re
import json
import os
from typing import List, Dict, Optional, Any
from datetime import datetime

class FinalPDFExtractor:
    """最終版 PDF 抽取器 - 完整功能版本"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.orders = []
        
        # 支援的材料代碼格式
        self.valid_patterns = [
            r'^H[A-Z]-',               # 任何 H?- 開頭 (HC-, HD-, HS-, HN-, HA-, HE-, HP-, HB-等)
            r'^I[AB][A-Z]{2,3}\d+z$',  # 特殊I系列: IAAD...z, IBAZ...z
            r'^[A-Z]{1,2}$',           # 簡化代碼: g, C
            r'^\d+$'                   # 數字代碼: 21
        ]
        
    def extract_orders(self) -> List[Dict[str, Any]]:
        """主要抽取函數"""
        print(f"🔍 開始處理 PDF: {self.pdf_path}")
        
        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                print(f"  處理第 {page_num + 1} 頁")
                
                text = page.extract_text()
                if text:
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    page_orders = self._parse_variable_format(lines)
                    self.orders.extend(page_orders)
        
        print(f"✅ 共抽取到 {len(self.orders)} 筆訂單")
        return self.orders
    
    def _parse_variable_format(self, lines: List[str]) -> List[Dict[str, Any]]:
        """解析可變格式資料（以PD開頭劃分區塊）"""
        orders = []
        i = 0
        
        while i < len(lines):
            if lines[i].startswith('PD'):
                start = i
                # 找到下一筆 PD 或文件結尾
                i += 1
                while i < len(lines) and not lines[i].startswith('PD'):
                    i += 1
                
                # 取得這筆訂單的所有行
                block_lines = lines[start:i]
                order = self._parse_order_block(block_lines)
                if order:
                    orders.append(order)
                    customer = order.get('客戶名稱', 'Unknown')
                    product = order.get('上階品名', 'Unknown')
                    material_count = len(order.get('耗料', []))
                    print(f"    ✅ {order['工單單號']} - {customer} - {product} ({material_count}種材料)")
            else:
                i += 1
        
        return orders
    
    def _parse_order_block(self, block_lines: List[str]) -> Optional[Dict[str, Any]]:
        """解析單個訂單區塊（可變行數）"""
        if len(block_lines) < 2:
            return None
        
        order = self._create_empty_order()
        
        # 第1行一定是主要資料
        if not self._parse_main_line(order, block_lines[0]):
            return None
        
        # 其餘行：次要資料或耗料
        for line in block_lines[1:]:
            if 'SD' in line or 'SA' in line:  # 包含訂單號的行
                self._parse_secondary_line(order, line)
            elif '耗料代碼' in line or '需求量' in line or '已領量' in line:
                # 跳過表頭行
                continue
            else:
                # 嘗試解析為材料行
                self._parse_material_line(order, line)
        
        # 最終處理
        self._finalize_order(order)
        
        return order
    
    def _parse_main_line(self, order: Dict[str, Any], line: str) -> bool:
        """解析主要資料行（放寬欄位要求）"""
        parts = line.split()
        if len(parts) < 6:  # 放寬要求：最少需要6個欄位
            return False
        
        try:
            order["工單單號"] = parts[0]  # PD20250805002
            order["上線日"] = parts[1].replace('/', '-')  # 2025/08/06 -> 2025-08-06
            order["客戶名稱"] = parts[2]  # 客戶名稱
            
            # 品號+品名處理
            no_and_name = parts[3]
            if '包膠' in no_and_name:
                order["NO"] = no_and_name.replace('包膠', '')
                order["品名"] = '包膠'
            elif '包套管' in no_and_name:
                order["NO"] = no_and_name.replace('包套管', '')
                order["品名"] = '包套管'
            elif '面層包膠' in no_and_name:
                order["NO"] = no_and_name.replace('面層包膠', '')
                order["品名"] = '面層包膠'
            else:
                order["NO"] = no_and_name
            
            order["上階品名"] = parts[4]
            order["上階規格"] = parts[5] if len(parts) > 5 else ""
            
            # 數字欄位處理（容錯處理）
            order["總長"] = 0
            order["數量"] = 1
            order["顏色"] = ""
            order["硬度"] = None
            order["硬度公差"] = None
            
            # 如果有足夠欄位，嘗試解析數字和顏色
            if len(parts) > 6:
                try:
                    order["總長"] = int(parts[6]) if parts[6].isdigit() else 0
                except (ValueError, IndexError):
                    pass
            
            if len(parts) > 7:
                try:
                    order["數量"] = int(parts[7]) if parts[7].isdigit() else 1
                except (ValueError, IndexError):
                    pass
            
            if len(parts) > 8:
                order["顏色"] = parts[8]
            
            # 硬度±公差處理（如果存在）
            if len(parts) > 9:
                try:
                    hardness_str = parts[9]
                    hardness_match = re.match(r'(\d+)±(\d+)', hardness_str)
                    if hardness_match:
                        order["硬度"] = int(hardness_match.group(1))
                        order["硬度公差"] = int(hardness_match.group(2))
                    elif hardness_str.isdigit():
                        order["硬度"] = int(hardness_str)
                        order["硬度公差"] = 5  # 預設公差
                except (ValueError, IndexError):
                    pass
            
            return True
            
        except (ValueError, IndexError) as e:
            print(f"      ⚠️ 主行解析錯誤: {e}")
            return False
    
    def _parse_secondary_line(self, order: Dict[str, Any], line: str):
        """解析次要資訊行"""
        parts = line.split()
        if len(parts) >= 2:
            order["產品類別"] = parts[0]  # A, B, C, etc.
            order["訂單單號"] = parts[1]  # SD20250804004-001
            
            # 備註處理（跳過固定文字）
            skip_words = {'耗料代碼', '需求量', '已領量', '模具'}
            remarks = []
            
            for part in parts[2:]:
                if part not in skip_words:
                    remarks.append(part)
            
            if remarks:
                order["客戶備註"] = ' '.join(remarks)
    
    def _parse_material_line(self, order: Dict[str, Any], line: str):
        """解析材料資訊行，支援同一行多組材料（追加模式）"""
        tokens = line.split()
        if len(tokens) < 2:
            return
        
        i = 0
        while i + 1 < len(tokens):
            code = tokens[i]
            
            # 檢查是否像材料代碼
            if any(re.match(pattern, code) for pattern in self.valid_patterns):
                try:
                    need_qty = float(tokens[i + 1])
                    received_qty = float(tokens[i + 2]) if i + 2 < len(tokens) else 0.0
                    
                    # 追加材料到訂單
                    material_item = {
                        "代碼": code,
                        "需求量": need_qty,
                        "已領量": received_qty
                    }
                    order.setdefault("耗料", []).append(material_item)
                    
                    # 特殊代碼提示
                    if re.match(r'^I[AB][A-Z]{2,3}\d+z$', code):
                        print(f"      🔍 特殊材料: {code}")
                    
                    # 跳過已處理的3個token（代碼、需求量、已領量）
                    i += 3
                    continue
                    
                except (ValueError, IndexError):
                    # 如果數量解析失敗，跳過這個token繼續
                    pass
            
            i += 1
    
    
    def _finalize_order(self, order: Dict[str, Any]):
        """最終處理和驗證訂單"""
        # 確保耗料欄位存在
        if "耗料" not in order:
            order["耗料"] = []
    
    def _create_empty_order(self) -> Dict[str, Any]:
        """創建空訂單結構"""
        return {
            "上線日": None,
            "客戶名稱": None,
            "上階品名": None,
            "上階規格": None,
            "數量": None,
            "硬度": None,
            "硬度公差": None,
            "客戶備註": None,
            "顏色": None,
            "工單單號": None,
            "產品類別": None,
            "訂單單號": None,
            "NO": None,
            "品名": None,
            "總長": None,
            "耗料": []
        }
    
    def save_results(self, output_dir: str = "output"):
        """儲存結果到多種格式"""
        os.makedirs(output_dir, exist_ok=True)
        
        base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 1. 儲存 JSON
        json_path = os.path.join(output_dir, f"{base_name}_extracted_{timestamp}.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.orders, f, ensure_ascii=False, indent=2)
        print(f"💾 JSON 已儲存: {json_path}")
        
        # 2. 儲存 Excel
        excel_path = os.path.join(output_dir, f"{base_name}_extracted_{timestamp}.xlsx")
        self._save_to_excel(excel_path)
        print(f"📊 Excel 已儲存: {excel_path}")
        
        return {"json": json_path, "excel": excel_path}
    
    def _save_to_excel(self, excel_path: str):
        """儲存到 Excel（多工作表，材料代碼分類）"""
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # 主訂單資料
            df_data = []
            for order in self.orders:
                row = order.copy()
                
                # 分類材料代碼和數量
                h_codes = []      # H系列代碼
                h_quantities = [] # H系列原料公斤數
                i_codes = []      # I系列代碼  
                i_quantities = [] # I系列鐵材隻數
                other_materials = []  # 其他材料
                
                if order.get("耗料"):
                    for material in order["耗料"]:
                        code = material.get("代碼", "")
                        qty = material.get("需求量", 0)
                        
                        if code.startswith('H'):
                            h_codes.append(code)
                            h_quantities.append(str(qty))  # 純數字，方便運算
                        elif code.startswith('I'):
                            i_codes.append(code)
                            i_quantities.append(str(qty))  # 純數字，方便運算
                        else:
                            other_materials.append(f"{code}({qty})")
                
                # 新增分類欄位
                row["H系列代碼"] = "; ".join(h_codes) if h_codes else ""
                row["原料公斤數"] = "; ".join(h_quantities) if h_quantities else ""
                row["I系列代碼"] = "; ".join(i_codes) if i_codes else ""
                row["鐵材隻數"] = "; ".join(i_quantities) if i_quantities else ""
                row["其他材料"] = "; ".join(other_materials) if other_materials else ""
                
                if "耗料" in row:
                    del row["耗料"]
                
                df_data.append(row)
            
            df = pd.DataFrame(df_data)
            
            # 重新排序欄位，將分類材料放在前面
            base_columns = ["工單單號", "訂單單號", "客戶名稱", "上階品名", "上階規格", 
                          "數量", "硬度", "硬度公差", "顏色", "上線日", "產品類別"]
            material_columns = ["H系列代碼", "原料公斤數", "I系列代碼", "鐵材隻數", "其他材料"]
            other_columns = [col for col in df.columns 
                           if col not in base_columns + material_columns 
                           and not col.startswith("耗料")]  # 排除舊的耗料欄位
            
            column_order = base_columns + material_columns + other_columns
            df = df.reindex(columns=[col for col in column_order if col in df.columns])
            
            df.to_excel(writer, sheet_name='訂單明細', index=False)
            
            # 材料統計表
            material_stats = self._get_material_statistics()
            df_material_stats = pd.DataFrame(material_stats)
            df_material_stats.to_excel(writer, sheet_name='材料統計', index=False)
            
            # 統計資料
            stats = self.get_statistics()
            df_stats = pd.DataFrame([stats])
            df_stats.to_excel(writer, sheet_name='統計摘要', index=False)
    
    def get_statistics(self) -> Dict[str, Any]:
        """獲取統計資料"""
        stats = {
            "總訂單數": len(self.orders),
            "客戶數量": len(set(o.get("客戶名稱") for o in self.orders if o.get("客戶名稱"))),
            "產品類型數": len(set(o.get("上階品名") for o in self.orders if o.get("上階品名"))),
            "總材料項目": sum(len(o.get("耗料", [])) for o in self.orders),
            "總需求量": sum(m.get("需求量", 0) for o in self.orders for m in o.get("耗料", [])),
            "處理時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        return stats
    
    def _get_material_statistics(self) -> List[Dict[str, Any]]:
        """獲取材料統計資料（按類別分組）"""
        h_materials = {}
        i_materials = {}
        other_materials = {}
        
        for order in self.orders:
            for material in order.get("耗料", []):
                code = material.get("代碼", "")
                qty = material.get("需求量", 0)
                
                if code.startswith('H'):
                    h_materials[code] = h_materials.get(code, 0) + qty
                elif code.startswith('I'):
                    i_materials[code] = i_materials.get(code, 0) + qty
                else:
                    other_materials[code] = other_materials.get(code, 0) + qty
        
        stats = []
        
        # H系列材料統計
        for code, total_qty in sorted(h_materials.items(), key=lambda x: x[1], reverse=True):
            stats.append({
                "材料類別": "H系列",
                "材料代碼": code,
                "總需求量": total_qty,
                "使用次數": sum(1 for order in self.orders 
                               for material in order.get("耗料", []) 
                               if material.get("代碼") == code)
            })
        
        # I系列材料統計
        for code, total_qty in sorted(i_materials.items(), key=lambda x: x[1], reverse=True):
            stats.append({
                "材料類別": "I系列",
                "材料代碼": code,
                "總需求量": total_qty,
                "使用次數": sum(1 for order in self.orders 
                               for material in order.get("耗料", []) 
                               if material.get("代碼") == code)
            })
        
        # 其他材料統計
        for code, total_qty in sorted(other_materials.items(), key=lambda x: x[1], reverse=True):
            stats.append({
                "材料類別": "其他",
                "材料代碼": code,
                "總需求量": total_qty,
                "使用次數": sum(1 for order in self.orders 
                               for material in order.get("耗料", []) 
                               if material.get("代碼") == code)
            })
        
        return stats
    
    def print_summary(self):
        """列印處理摘要"""
        stats = self.get_statistics()
        
        print(f"\n{'='*60}")
        print(f"📋 PDF 抽取摘要報告")
        print(f"{'='*60}")
        print(f"📁 檔案: {os.path.basename(self.pdf_path)}")
        print(f"📊 總訂單數: {stats['總訂單數']}")
        print(f"👥 客戶數量: {stats['客戶數量']}")
        print(f"🔧 產品類型: {stats['產品類型數']}")
        print(f"🧪 材料項目: {stats['總材料項目']}")
        print(f"⚖️  總需求量: {stats['總需求量']:.1f} kg")
        print(f"⏰ 處理時間: {stats['處理時間']}")
        
        # 客戶分布
        customers = {}
        materials = {}
        
        for order in self.orders:
            customer = order.get("客戶名稱", "未知")
            customers[customer] = customers.get(customer, 0) + 1
            
            for material in order.get("耗料", []):
                code = material.get("代碼", "")
                if code:
                    materials[code] = materials.get(code, 0) + material.get("需求量", 0)
        
        print(f"\n👥 主要客戶 (前5名):")
        for customer, count in sorted(customers.items(), key=lambda x: x[1], reverse=True)[:5]:
            print(f"   {customer}: {count} 筆")
        
        print(f"\n🧪 主要材料 (前5種):")
        for material, qty in sorted(materials.items(), key=lambda x: x[1], reverse=True)[:5]:
            print(f"   {material}: {qty:.1f} kg")


def main():
    """主程式"""
    import sys
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("請輸入 PDF 檔案路徑: ").strip()
    
    if not pdf_path or not os.path.exists(pdf_path):
        print("❌ 錯誤: 檔案不存在")
        return
    
    # 建立抽取器
    extractor = FinalPDFExtractor(pdf_path)
    
    # 抽取訂單
    orders = extractor.extract_orders()
    
    if orders:
        # 顯示摘要
        extractor.print_summary()
        
        # 儲存結果
        saved_files = extractor.save_results()
        
        print(f"\n🎉 處理完成！")
        print(f"📁 檔案已儲存到: {list(saved_files.values())}")
        
        # 顯示第一筆完整資料
        if orders:
            print(f"\n📄 第一筆訂單範例:")
            print(json.dumps(orders[0], ensure_ascii=False, indent=2))
    else:
        print("❌ 未能抽取到任何訂單")


if __name__ == "__main__":
    main()