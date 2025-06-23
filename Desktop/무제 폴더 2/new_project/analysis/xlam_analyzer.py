#!/usr/bin/env python3
"""
Excel Add-in (.xlam) 파일 분석기
DSD Breaker의 기능과 구조를 분석합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
from pathlib import Path
import json

class XlamAnalyzer:
    """Excel Add-in 분석 클래스"""
    
    def __init__(self, xlam_path):
        self.xlam_path = Path(xlam_path)
        self.temp_dir = Path("temp_analysis")
        self.analysis_result = {}
        
    def extract_xlam(self):
        """XLAM 파일을 압축 해제하여 내부 구조 분석"""
        try:
            # 임시 폴더 생성
            self.temp_dir.mkdir(exist_ok=True)
            
            # XLAM 파일은 ZIP 형태로 압축된 파일
            with zipfile.ZipFile(self.xlam_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
            
            print(f"✅ XLAM 파일 압축 해제 완료: {self.temp_dir}")
            return True
            
        except Exception as e:
            print(f"❌ 압축 해제 실패: {e}")
            return False
    
    def analyze_structure(self):
        """파일 구조 분석"""
        if not self.temp_dir.exists():
            print("❌ 압축 해제된 파일이 없습니다.")
            return
        
        structure = {}
        
        for root, dirs, files in os.walk(self.temp_dir):
            rel_path = os.path.relpath(root, self.temp_dir)
            if rel_path == '.':
                rel_path = 'root'
            
            structure[rel_path] = {
                'directories': dirs,
                'files': files
            }
        
        self.analysis_result['structure'] = structure
        
        print("📁 파일 구조:")
        for path, content in structure.items():
            print(f"  {path}/")
            for file in content['files']:
                print(f"    📄 {file}")
        
        return structure
    
    def analyze_vba_modules(self):
        """VBA 모듈 분석"""
        vba_dir = self.temp_dir / "xl" / "vbaProject.bin"
        
        if vba_dir.exists():
            print("🔍 VBA 프로젝트 발견")
            # VBA 코드는 바이너리 형태로 저장되어 직접 분석이 어려움
            # 파일 크기와 존재 여부만 확인
            file_size = vba_dir.stat().st_size
            self.analysis_result['vba'] = {
                'exists': True,
                'size': file_size,
                'note': 'VBA 코드는 바이너리 형태로 저장됨'
            }
            print(f"  📊 VBA 프로젝트 크기: {file_size:,} bytes")
        else:
            self.analysis_result['vba'] = {'exists': False}
            print("❌ VBA 프로젝트 없음")
    
    def analyze_xml_files(self):
        """XML 파일들 분석"""
        xml_info = {}
        
        # 주요 XML 파일들 분석
        xml_files = [
            "xl/workbook.xml",
            "xl/app.xml", 
            "docProps/core.xml",
            "docProps/app.xml"
        ]
        
        for xml_file in xml_files:
            xml_path = self.temp_dir / xml_file
            if xml_path.exists():
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    
                    xml_info[xml_file] = {
                        'root_tag': root.tag,
                        'namespace': root.tag.split('}')[0][1:] if '}' in root.tag else '',
                        'children_count': len(list(root)),
                        'attributes': dict(root.attrib)
                    }
                    
                    print(f"📄 {xml_file}: {root.tag}")
                    
                except Exception as e:
                    xml_info[xml_file] = {'error': str(e)}
                    print(f"❌ {xml_file} 분석 실패: {e}")
        
        self.analysis_result['xml_files'] = xml_info
        return xml_info
    
    def analyze_relationships(self):
        """관계 파일 분석"""
        rels_dir = self.temp_dir / "_rels"
        xl_rels_dir = self.temp_dir / "xl" / "_rels"
        
        relationships = {}
        
        for rels_path in [rels_dir / ".rels", xl_rels_dir / "workbook.xml.rels"]:
            if rels_path.exists():
                try:
                    tree = ET.parse(rels_path)
                    root = tree.getroot()
                    
                    rels = []
                    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rels.append({
                            'id': rel.get('Id'),
                            'type': rel.get('Type'),
                            'target': rel.get('Target')
                        })
                    
                    relationships[str(rels_path.relative_to(self.temp_dir))] = rels
                    print(f"🔗 {rels_path.name}: {len(rels)}개 관계")
                    
                except Exception as e:
                    print(f"❌ {rels_path} 분석 실패: {e}")
        
        self.analysis_result['relationships'] = relationships
        return relationships
    
    def detect_add_in_features(self):
        """Add-in 특징 탐지"""
        features = []
        
        # VBA 프로젝트 존재 여부
        if self.analysis_result.get('vba', {}).get('exists'):
            features.append("VBA 매크로 포함")
        
        # 특정 XML 파일들 확인
        if 'xml_files' in self.analysis_result:
            if 'xl/workbook.xml' in self.analysis_result['xml_files']:
                features.append("워크북 구조 포함")
        
        # 관계 파일에서 Add-in 특징 확인
        if 'relationships' in self.analysis_result:
            for rel_file, rels in self.analysis_result['relationships'].items():
                for rel in rels:
                    if 'addin' in rel.get('type', '').lower():
                        features.append("Add-in 관계 정의")
                    if 'vba' in rel.get('target', '').lower():
                        features.append("VBA 프로젝트 연결")
        
        self.analysis_result['features'] = features
        
        print("\n🎯 탐지된 Add-in 특징:")
        for feature in features:
            print(f"  ✅ {feature}")
        
        return features
    
    def save_analysis_report(self, output_path="analysis_report.json"):
        """분석 결과를 JSON 파일로 저장"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.analysis_result, f, indent=2, ensure_ascii=False)
            
            print(f"\n📊 분석 보고서 저장: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 보고서 저장 실패: {e}")
            return False
    
    def cleanup(self):
        """임시 파일 정리"""
        import shutil
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            print(f"🧹 임시 파일 정리 완료")
    
    def full_analysis(self):
        """전체 분석 실행"""
        print("🔍 DSD Breaker XLAM 파일 분석 시작\n")
        
        # 1. 압축 해제
        if not self.extract_xlam():
            return False
        
        # 2. 구조 분석
        print("\n" + "="*50)
        print("📁 파일 구조 분석")
        print("="*50)
        self.analyze_structure()
        
        # 3. VBA 분석
        print("\n" + "="*50)
        print("🔧 VBA 모듈 분석")
        print("="*50)
        self.analyze_vba_modules()
        
        # 4. XML 분석
        print("\n" + "="*50)
        print("📄 XML 파일 분석")
        print("="*50)
        self.analyze_xml_files()
        
        # 5. 관계 분석
        print("\n" + "="*50)
        print("🔗 관계 파일 분석")
        print("="*50)
        self.analyze_relationships()
        
        # 6. 특징 탐지
        print("\n" + "="*50)
        print("🎯 Add-in 특징 탐지")
        print("="*50)
        self.detect_add_in_features()
        
        # 7. 보고서 저장
        self.save_analysis_report("analysis/dsd_breaker_analysis.json")
        
        # 8. 정리
        self.cleanup()
        
        print("\n✅ 분석 완료!")
        return True

def main():
    """메인 실행 함수"""
    xlam_file = "../original/DsDbreaker.xlam"
    
    if not os.path.exists(xlam_file):
        print(f"❌ 파일을 찾을 수 없습니다: {xlam_file}")
        return
    
    analyzer = XlamAnalyzer(xlam_file)
    analyzer.full_analysis()

if __name__ == "__main__":
    main()