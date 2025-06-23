#!/usr/bin/env python3
"""
DSD Breaker 분석 스크립트 (경로 수정 버전)
"""

import zipfile
import xml.etree.ElementTree as ET
import os
from pathlib import Path
import json
import shutil

def analyze_dsd_breaker():
    """DSD Breaker XLAM 파일 분석"""
    
    xlam_path = "/Users/youngjunlee/Desktop/무제 폴더 2/new_project/original/DsDbreaker.xlam"
    temp_dir = Path("temp_analysis")
    analysis_result = {}
    
    print("🔍 DSD Breaker XLAM 파일 분석 시작")
    print(f"📁 분석 대상: {xlam_path}")
    
    if not os.path.exists(xlam_path):
        print(f"❌ 파일을 찾을 수 없습니다: {xlam_path}")
        return
    
    try:
        # 1. 압축 해제
        print("\n🔄 XLAM 파일 압축 해제 중...")
        temp_dir.mkdir(exist_ok=True)
        
        with zipfile.ZipFile(xlam_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        print("✅ 압축 해제 완료")
        
        # 2. 파일 구조 분석
        print("\n📁 파일 구조 분석:")
        structure = {}
        
        for root, dirs, files in os.walk(temp_dir):
            rel_path = os.path.relpath(root, temp_dir)
            if rel_path == '.':
                rel_path = 'root'
            
            structure[rel_path] = files
            print(f"  📂 {rel_path}/")
            for file in files:
                print(f"    📄 {file}")
        
        analysis_result['structure'] = structure
        
        # 3. VBA 프로젝트 확인
        print("\n🔧 VBA 프로젝트 분석:")
        vba_path = temp_dir / "xl" / "vbaProject.bin"
        
        if vba_path.exists():
            file_size = vba_path.stat().st_size
            print(f"  ✅ VBA 프로젝트 발견 (크기: {file_size:,} bytes)")
            analysis_result['vba'] = {'exists': True, 'size': file_size}
        else:
            print("  ❌ VBA 프로젝트 없음")
            analysis_result['vba'] = {'exists': False}
        
        # 4. 주요 XML 파일 분석
        print("\n📄 XML 파일 분석:")
        xml_files = [
            "xl/workbook.xml",
            "docProps/core.xml",
            "docProps/app.xml"
        ]
        
        xml_info = {}
        for xml_file in xml_files:
            xml_path = temp_dir / xml_file
            if xml_path.exists():
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    
                    # XML 내용 간단 분석
                    print(f"  ✅ {xml_file}")
                    print(f"    루트 태그: {root.tag}")
                    
                    # 특별한 정보 추출
                    if xml_file == "docProps/app.xml":
                        # 애플리케이션 정보
                        for elem in root.iter():
                            if 'Application' in elem.tag and elem.text:
                                print(f"    애플리케이션: {elem.text}")
                            elif 'Company' in elem.tag and elem.text:
                                print(f"    회사: {elem.text}")
                    
                    xml_info[xml_file] = {'root_tag': root.tag, 'exists': True}
                    
                except Exception as e:
                    print(f"  ❌ {xml_file} 분석 실패: {e}")
                    xml_info[xml_file] = {'exists': True, 'error': str(e)}
            else:
                print(f"  ❌ {xml_file} 없음")
                xml_info[xml_file] = {'exists': False}
        
        analysis_result['xml_files'] = xml_info
        
        # 5. 관계 파일 분석
        print("\n🔗 관계 파일 분석:")
        rels_files = [
            "_rels/.rels",
            "xl/_rels/workbook.xml.rels"
        ]
        
        relationships = {}
        for rels_file in rels_files:
            rels_path = temp_dir / rels_file
            if rels_path.exists():
                try:
                    tree = ET.parse(rels_path)
                    root = tree.getroot()
                    
                    rels = []
                    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rel_info = {
                            'id': rel.get('Id'),
                            'type': rel.get('Type', '').split('/')[-1],  # 타입의 마지막 부분만
                            'target': rel.get('Target')
                        }
                        rels.append(rel_info)
                    
                    relationships[rels_file] = rels
                    print(f"  ✅ {rels_file}: {len(rels)}개 관계")
                    
                    for rel in rels:
                        print(f"    - {rel['type']}: {rel['target']}")
                        
                except Exception as e:
                    print(f"  ❌ {rels_file} 분석 실패: {e}")
            else:
                print(f"  ❌ {rels_file} 없음")
        
        analysis_result['relationships'] = relationships
        
        # 6. Add-in 특징 탐지
        print("\n🎯 Add-in 특징 탐지:")
        features = []
        
        if analysis_result['vba']['exists']:
            features.append("VBA 매크로 포함")
            print("  ✅ VBA 매크로 포함")
        
        # 파일 확장자가 .xlam인 경우
        features.append("Excel Add-in 형식")
        print("  ✅ Excel Add-in 형식 (.xlam)")
        
        # 특정 관계 확인
        for rels_file, rels in relationships.items():
            for rel in rels:
                if 'macro' in rel['type'].lower() or 'vba' in rel['target'].lower():
                    features.append("매크로 관계 정의")
                    print("  ✅ 매크로 관계 정의")
                    break
        
        analysis_result['features'] = features
        
        # 7. 결과 저장
        output_file = "/Users/youngjunlee/Desktop/무제 폴더 2/new_project/analysis/dsd_breaker_analysis.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(analysis_result, f, indent=2, ensure_ascii=False)
        
        print(f"\n📊 분석 결과 저장: {output_file}")
        
        # 8. 요약 출력
        print("\n" + "="*60)
        print("📋 DSD BREAKER 분석 요약")
        print("="*60)
        print(f"파일 크기: {os.path.getsize(xlam_path):,} bytes")
        print(f"VBA 프로젝트: {'✅ 있음' if analysis_result['vba']['exists'] else '❌ 없음'}")
        print(f"주요 특징: {', '.join(features)}")
        print(f"파일 구조: {len(structure)}개 폴더")
        
        total_files = sum(len(files) for files in structure.values())
        print(f"총 파일 수: {total_files}개")
        
    except Exception as e:
        print(f"❌ 분석 중 오류 발생: {e}")
        
    finally:
        # 임시 파일 정리
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
            print("\n🧹 임시 파일 정리 완료")
        
        print("\n✅ DSD Breaker 분석 완료!")

if __name__ == "__main__":
    analyze_dsd_breaker()