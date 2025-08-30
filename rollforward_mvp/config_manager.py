"""
설정 관리 시스템 - 보안 강화 버전
민감한 정보 보호, 유효성 검증, 복구 메커니즘 포함

보안 고려사항:
- 설정 파일 경로 검증
- JSON 스키마 기반 유효성 검사
- 민감한 정보 암호화 (필요시)
- 기본값 설정 및 자동 복구
- 로그 보안 (민감한 정보 마스킹)
"""

import json
import os
import logging
import shutil
from pathlib import Path
from typing import Dict, List, Any, Optional, Union
from dataclasses import dataclass, asdict
from datetime import datetime
import hashlib
from jsonschema import validate, ValidationError, Draft7Validator


@dataclass
class WorksheetProcessingRule:
    """워크시트 처리 규칙 데이터 구조"""
    worksheet_pattern: str
    processing_strategy: str  # 'table_level', 'worksheet_level', 'user_choice'
    confidence_threshold: float
    auto_approve: bool
    comment: str = ""


@dataclass
class SecurityConfig:
    """보안 설정 구조"""
    max_file_size_mb: int = 100
    allowed_extensions: List[str] = None
    backup_enabled: bool = True
    log_sensitive_data: bool = False
    config_encryption: bool = False
    
    def __post_init__(self):
        if self.allowed_extensions is None:
            self.allowed_extensions = ['.xlsx', '.xlsm', '.xls']


class ConfigurationError(Exception):
    """설정 관련 예외 클래스"""
    pass


class SecurityValidationError(Exception):
    """보안 검증 실패 예외 클래스"""
    pass


class ConfigManager:
    """보안 강화된 설정 관리 클래스"""
    
    # 설정 파일 스키마 정의 (보안 검증용)
    CONFIG_SCHEMA = {
        "type": "object",
        "properties": {
            "version": {"type": "string"},
            "general": {
                "type": "object",
                "properties": {
                    "auto_backup": {"type": "boolean"},
                    "log_level": {"type": "string", "enum": ["DEBUG", "INFO", "WARNING", "ERROR"]},
                    "max_concurrent_files": {"type": "integer", "minimum": 1, "maximum": 10},
                    "default_processing_timeout": {"type": "integer", "minimum": 30}
                }
            },
            "processing": {
                "type": "object",
                "properties": {
                    "default_strategy": {"type": "string", "enum": ["table_level", "worksheet_level", "user_choice"]},
                    "confidence_threshold": {"type": "number", "minimum": 0.0, "maximum": 100.0},
                    "auto_approve_high_confidence": {"type": "boolean"}
                }
            },
            "security": {
                "type": "object",
                "properties": {
                    "max_file_size_mb": {"type": "integer", "minimum": 1, "maximum": 1000},
                    "allowed_extensions": {"type": "array", "items": {"type": "string"}},
                    "backup_enabled": {"type": "boolean"},
                    "log_sensitive_data": {"type": "boolean"}
                }
            }
        },
        "required": ["version", "general", "processing", "security"],
        "additionalProperties": False
    }
    
    WHITELIST_SCHEMA = {
        "type": "object",
        "properties": {
            "version": {"type": "string"},
            "worksheet_patterns": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "pattern": {"type": "string"},
                        "strategy": {"type": "string", "enum": ["table_level", "worksheet_level"]},
                        "auto_approve": {"type": "boolean"},
                        "comment": {"type": "string"}
                    },
                    "required": ["pattern", "strategy"],
                    "additionalProperties": False
                }
            }
        },
        "required": ["version", "worksheet_patterns"],
        "additionalProperties": False
    }
    
    BLACKLIST_SCHEMA = {
        "type": "object",
        "properties": {
            "version": {"type": "string"},
            "excluded_patterns": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "pattern": {"type": "string"},
                        "reason": {"type": "string"},
                        "exclude_completely": {"type": "boolean"}
                    },
                    "required": ["pattern", "reason"],
                    "additionalProperties": False
                }
            }
        },
        "required": ["version", "excluded_patterns"],
        "additionalProperties": False
    }
    
    def __init__(self, config_dir: Optional[str] = None):
        """
        설정 관리자 초기화
        
        Args:
            config_dir: 설정 파일 디렉토리 경로 (None이면 현재 디렉토리/config)
        """
        # 설정 디렉토리 보안 검증
        self.config_dir = Path(config_dir) if config_dir else Path.cwd() / 'config'
        self._validate_config_directory()
        
        # 설정 파일 경로
        self.config_file = self.config_dir / 'rollforward_config.json'
        self.whitelist_file = self.config_dir / 'whitelist.json'
        self.blacklist_file = self.config_dir / 'blacklist.json'
        
        # 로깅 설정
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        # 설정 캐시
        self._config_cache = {}
        self._cache_timestamp = {}
        
        # 보안 설정
        self.security_config = SecurityConfig()
        
        # 초기화
        self._ensure_config_directory()
        self._initialize_default_configs()
    
    def _validate_config_directory(self):
        """설정 디렉토리 보안 검증"""
        if not isinstance(self.config_dir, Path):
            raise SecurityValidationError("설정 디렉토리 경로가 유효하지 않습니다")
        
        # 상위 디렉토리 경로 조작 방지
        try:
            self.config_dir = self.config_dir.resolve()
            if '..' in str(self.config_dir):
                raise SecurityValidationError("상위 디렉토리 접근은 허용되지 않습니다")
        except Exception as e:
            raise SecurityValidationError(f"디렉토리 경로 검증 실패: {e}")
    
    def _ensure_config_directory(self):
        """설정 디렉토리 생성 및 권한 설정"""
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            
            # 윈도우가 아닌 경우 디렉토리 권한 설정
            if os.name != 'nt':
                os.chmod(self.config_dir, 0o700)  # 소유자만 읽기/쓰기/실행
                
        except PermissionError:
            raise ConfigurationError(f"설정 디렉토리 생성 권한 없음: {self.config_dir}")
        except Exception as e:
            raise ConfigurationError(f"설정 디렉토리 생성 실패: {e}")
    
    def _initialize_default_configs(self):
        """기본 설정 파일들 초기화"""
        # 메인 설정
        if not self.config_file.exists():
            self._create_default_config()
        
        # 화이트리스트
        if not self.whitelist_file.exists():
            self._create_default_whitelist()
        
        # 블랙리스트
        if not self.blacklist_file.exists():
            self._create_default_blacklist()
    
    def _create_default_config(self):
        """기본 메인 설정 생성"""
        default_config = {
            "version": "1.0.0",
            "general": {
                "auto_backup": True,
                "log_level": "INFO",
                "max_concurrent_files": 3,
                "default_processing_timeout": 300
            },
            "processing": {
                "default_strategy": "user_choice",
                "confidence_threshold": 70.0,
                "auto_approve_high_confidence": False
            },
            "security": {
                "max_file_size_mb": 100,
                "allowed_extensions": [".xlsx", ".xlsm", ".xls"],
                "backup_enabled": True,
                "log_sensitive_data": False
            }
        }
        
        self._save_config_secure(self.config_file, default_config, self.CONFIG_SCHEMA)
        self.logger.info("기본 설정 파일이 생성되었습니다")
    
    def _create_default_whitelist(self):
        """기본 화이트리스트 생성"""
        default_whitelist = {
            "version": "1.0.0",
            "worksheet_patterns": [
                {
                    "pattern": "별도.*BS|별도.*재무상태표",
                    "strategy": "worksheet_level",
                    "auto_approve": True,
                    "comment": "별도 재무상태표는 항상 워크시트 전체 복사"
                },
                {
                    "pattern": "현금.*흐름표|CF|cash.*flow",
                    "strategy": "worksheet_level", 
                    "auto_approve": True,
                    "comment": "현금흐름표는 워크시트 전체 복사 권장"
                },
                {
                    "pattern": "메인.*BS|main.*balance",
                    "strategy": "table_level",
                    "auto_approve": True,
                    "comment": "메인 재무상태표는 테이블별 정밀 처리"
                }
            ]
        }
        
        self._save_config_secure(self.whitelist_file, default_whitelist, self.WHITELIST_SCHEMA)
        self.logger.info("기본 화이트리스트가 생성되었습니다")
    
    def _create_default_blacklist(self):
        """기본 블랙리스트 생성"""
        default_blacklist = {
            "version": "1.0.0",
            "excluded_patterns": [
                {
                    "pattern": "temp|임시|작업|test",
                    "reason": "임시 작업 파일로 처리 제외",
                    "exclude_completely": True
                },
                {
                    "pattern": "backup|백업|복사본",
                    "reason": "백업 파일로 처리 불필요",
                    "exclude_completely": True
                },
                {
                    "pattern": "log|로그|기록",
                    "reason": "로그 파일로 롤포워딩 대상 아님",
                    "exclude_completely": False
                }
            ]
        }
        
        self._save_config_secure(self.blacklist_file, default_blacklist, self.BLACKLIST_SCHEMA)
        self.logger.info("기본 블랙리스트가 생성되었습니다")
    
    def _save_config_secure(self, file_path: Path, data: Dict, schema: Dict):
        """보안을 고려한 설정 파일 저장"""
        try:
            # 스키마 유효성 검사
            validate(instance=data, schema=schema)
            
            # 임시 파일로 안전한 저장
            temp_file = file_path.with_suffix('.tmp')
            
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            # 원자적 이동 (atomic move)
            shutil.move(str(temp_file), str(file_path))
            
            # 윈도우가 아닌 경우 파일 권한 설정
            if os.name != 'nt':
                os.chmod(file_path, 0o600)  # 소유자만 읽기/쓰기
            
        except ValidationError as e:
            raise ConfigurationError(f"설정 유효성 검사 실패: {e.message}")
        except Exception as e:
            # 임시 파일 정리
            if temp_file.exists():
                temp_file.unlink()
            raise ConfigurationError(f"설정 저장 실패: {e}")
    
    def _load_config_secure(self, file_path: Path, schema: Dict) -> Dict:
        """보안을 고려한 설정 파일 로드"""
        try:
            # 파일 존재 및 권한 확인
            if not file_path.exists():
                raise ConfigurationError(f"설정 파일이 존재하지 않습니다: {file_path}")
            
            # 파일 크기 제한 (보안)
            file_size = file_path.stat().st_size
            if file_size > 1024 * 1024:  # 1MB 제한
                raise SecurityValidationError(f"설정 파일이 너무 큽니다: {file_size} bytes")
            
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 스키마 유효성 검사
            validate(instance=data, schema=schema)
            
            return data
            
        except json.JSONDecodeError as e:
            raise ConfigurationError(f"JSON 파싱 오류: {e}")
        except ValidationError as e:
            raise ConfigurationError(f"설정 유효성 검사 실패: {e.message}")
        except Exception as e:
            raise ConfigurationError(f"설정 로드 실패: {e}")
    
    def get_config(self, force_reload: bool = False) -> Dict:
        """메인 설정 로드 (캐싱 지원)"""
        cache_key = 'main_config'
        
        if not force_reload and self._is_cache_valid(cache_key, self.config_file):
            return self._config_cache[cache_key]
        
        config = self._load_config_secure(self.config_file, self.CONFIG_SCHEMA)
        self._update_cache(cache_key, config, self.config_file)
        
        self.logger.info("메인 설정이 로드되었습니다")
        return config
    
    def get_whitelist(self, force_reload: bool = False) -> List[WorksheetProcessingRule]:
        """화이트리스트 로드"""
        cache_key = 'whitelist'
        
        if not force_reload and self._is_cache_valid(cache_key, self.whitelist_file):
            return self._config_cache[cache_key]
        
        whitelist_data = self._load_config_secure(self.whitelist_file, self.WHITELIST_SCHEMA)
        
        # 데이터 구조 변환
        rules = []
        for item in whitelist_data['worksheet_patterns']:
            rule = WorksheetProcessingRule(
                worksheet_pattern=item['pattern'],
                processing_strategy=item['strategy'],
                confidence_threshold=70.0,  # 기본값
                auto_approve=item.get('auto_approve', False),
                comment=item.get('comment', '')
            )
            rules.append(rule)
        
        self._update_cache(cache_key, rules, self.whitelist_file)
        self.logger.info(f"화이트리스트 로드 완료: {len(rules)}개 규칙")
        return rules
    
    def get_blacklist(self, force_reload: bool = False) -> List[Dict]:
        """블랙리스트 로드"""
        cache_key = 'blacklist'
        
        if not force_reload and self._is_cache_valid(cache_key, self.blacklist_file):
            return self._config_cache[cache_key]
        
        blacklist_data = self._load_config_secure(self.blacklist_file, self.BLACKLIST_SCHEMA)
        patterns = blacklist_data['excluded_patterns']
        
        self._update_cache(cache_key, patterns, self.blacklist_file)
        self.logger.info(f"블랙리스트 로드 완료: {len(patterns)}개 패턴")
        return patterns
    
    def add_whitelist_rule(self, rule: WorksheetProcessingRule):
        """화이트리스트에 새 규칙 추가"""
        try:
            current_rules = self.get_whitelist(force_reload=True)
            
            # 중복 패턴 확인
            for existing_rule in current_rules:
                if existing_rule.worksheet_pattern == rule.worksheet_pattern:
                    raise ConfigurationError(f"이미 존재하는 패턴입니다: {rule.worksheet_pattern}")
            
            # 새 규칙 추가
            current_rules.append(rule)
            
            # 저장용 데이터 구조 변환
            whitelist_data = {
                "version": "1.0.0",
                "worksheet_patterns": [
                    {
                        "pattern": r.worksheet_pattern,
                        "strategy": r.processing_strategy,
                        "auto_approve": r.auto_approve,
                        "comment": r.comment
                    }
                    for r in current_rules
                ]
            }
            
            self._save_config_secure(self.whitelist_file, whitelist_data, self.WHITELIST_SCHEMA)
            
            # 캐시 무효화
            self._invalidate_cache('whitelist')
            
            self.logger.info(f"화이트리스트 규칙 추가: {rule.worksheet_pattern}")
            
        except Exception as e:
            self.logger.error(f"화이트리스트 규칙 추가 실패: {e}")
            raise ConfigurationError(f"규칙 추가 실패: {e}")
    
    def add_blacklist_pattern(self, pattern: str, reason: str, exclude_completely: bool = True):
        """블랙리스트에 새 패턴 추가"""
        try:
            blacklist_data = self._load_config_secure(self.blacklist_file, self.BLACKLIST_SCHEMA)
            
            # 중복 패턴 확인
            for existing_pattern in blacklist_data['excluded_patterns']:
                if existing_pattern['pattern'] == pattern:
                    raise ConfigurationError(f"이미 존재하는 패턴입니다: {pattern}")
            
            # 새 패턴 추가
            new_pattern = {
                "pattern": pattern,
                "reason": reason,
                "exclude_completely": exclude_completely
            }
            blacklist_data['excluded_patterns'].append(new_pattern)
            
            self._save_config_secure(self.blacklist_file, blacklist_data, self.BLACKLIST_SCHEMA)
            
            # 캐시 무효화
            self._invalidate_cache('blacklist')
            
            self.logger.info(f"블랙리스트 패턴 추가: {pattern}")
            
        except Exception as e:
            self.logger.error(f"블랙리스트 패턴 추가 실패: {e}")
            raise ConfigurationError(f"패턴 추가 실패: {e}")
    
    def update_config_value(self, section: str, key: str, value: Any):
        """설정값 업데이트"""
        try:
            config = self.get_config(force_reload=True)
            
            if section not in config:
                raise ConfigurationError(f"존재하지 않는 설정 섹션: {section}")
            
            # 보안 검증 - 민감한 설정 변경 제한
            if section == 'security' and key in ['max_file_size_mb', 'allowed_extensions']:
                if not self._validate_security_value(key, value):
                    raise SecurityValidationError(f"보안상 허용되지 않는 값: {key}={value}")
            
            config[section][key] = value
            
            self._save_config_secure(self.config_file, config, self.CONFIG_SCHEMA)
            
            # 캐시 무효화
            self._invalidate_cache('main_config')
            
            self.logger.info(f"설정 업데이트: {section}.{key} = {self._mask_sensitive_value(key, value)}")
            
        except Exception as e:
            self.logger.error(f"설정 업데이트 실패: {e}")
            raise ConfigurationError(f"설정 업데이트 실패: {e}")
    
    def _validate_security_value(self, key: str, value: Any) -> bool:
        """보안 설정 값 검증"""
        if key == 'max_file_size_mb':
            return isinstance(value, int) and 1 <= value <= 1000
        elif key == 'allowed_extensions':
            return (isinstance(value, list) and 
                   all(isinstance(ext, str) and ext.startswith('.') for ext in value))
        return True
    
    def _mask_sensitive_value(self, key: str, value: Any) -> str:
        """민감한 값 마스킹"""
        sensitive_keys = ['password', 'token', 'secret', 'key']
        if any(sensitive in key.lower() for sensitive in sensitive_keys):
            return '*' * 8
        return str(value)
    
    def _is_cache_valid(self, cache_key: str, file_path: Path) -> bool:
        """캐시 유효성 검사"""
        if cache_key not in self._config_cache:
            return False
        
        if cache_key not in self._cache_timestamp:
            return False
        
        try:
            file_mtime = file_path.stat().st_mtime
            return file_mtime <= self._cache_timestamp[cache_key]
        except:
            return False
    
    def _update_cache(self, cache_key: str, data: Any, file_path: Path):
        """캐시 업데이트"""
        self._config_cache[cache_key] = data
        try:
            self._cache_timestamp[cache_key] = file_path.stat().st_mtime
        except:
            self._cache_timestamp[cache_key] = datetime.now().timestamp()
    
    def _invalidate_cache(self, cache_key: str):
        """캐시 무효화"""
        if cache_key in self._config_cache:
            del self._config_cache[cache_key]
        if cache_key in self._cache_timestamp:
            del self._cache_timestamp[cache_key]
    
    def backup_configs(self) -> str:
        """설정 파일들 백업"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = self.config_dir / f'backup_{timestamp}'
            backup_dir.mkdir(exist_ok=True)
            
            # 설정 파일들 복사
            for config_file in [self.config_file, self.whitelist_file, self.blacklist_file]:
                if config_file.exists():
                    backup_file = backup_dir / config_file.name
                    shutil.copy2(config_file, backup_file)
            
            self.logger.info(f"설정 백업 완료: {backup_dir}")
            return str(backup_dir)
            
        except Exception as e:
            self.logger.error(f"설정 백업 실패: {e}")
            raise ConfigurationError(f"백업 실패: {e}")
    
    def restore_configs(self, backup_path: str):
        """설정 파일들 복원"""
        try:
            backup_dir = Path(backup_path)
            if not backup_dir.exists():
                raise ConfigurationError(f"백업 디렉토리가 존재하지 않습니다: {backup_path}")
            
            # 현재 설정 백업 (복원 실패 시 롤백용)
            rollback_backup = self.backup_configs()
            
            try:
                # 백업에서 복원
                for config_name in ['rollforward_config.json', 'whitelist.json', 'blacklist.json']:
                    backup_file = backup_dir / config_name
                    if backup_file.exists():
                        target_file = self.config_dir / config_name
                        shutil.copy2(backup_file, target_file)
                
                # 캐시 모두 무효화
                self._config_cache.clear()
                self._cache_timestamp.clear()
                
                self.logger.info(f"설정 복원 완료: {backup_path}")
                
            except Exception as e:
                # 복원 실패 시 롤백
                self.logger.warning(f"복원 실패, 롤백 수행: {e}")
                self.restore_configs(rollback_backup)
                raise
                
        except Exception as e:
            self.logger.error(f"설정 복원 실패: {e}")
            raise ConfigurationError(f"복원 실패: {e}")
    
    def validate_all_configs(self) -> Dict[str, bool]:
        """모든 설정 파일 유효성 검사"""
        results = {}
        
        # 메인 설정
        try:
            self._load_config_secure(self.config_file, self.CONFIG_SCHEMA)
            results['main_config'] = True
        except Exception as e:
            results['main_config'] = False
            self.logger.error(f"메인 설정 유효성 검사 실패: {e}")
        
        # 화이트리스트
        try:
            self._load_config_secure(self.whitelist_file, self.WHITELIST_SCHEMA)
            results['whitelist'] = True
        except Exception as e:
            results['whitelist'] = False
            self.logger.error(f"화이트리스트 유효성 검사 실패: {e}")
        
        # 블랙리스트
        try:
            self._load_config_secure(self.blacklist_file, self.BLACKLIST_SCHEMA)
            results['blacklist'] = True
        except Exception as e:
            results['blacklist'] = False
            self.logger.error(f"블랙리스트 유효성 검사 실패: {e}")
        
        return results
    
    def get_config_summary(self) -> Dict:
        """설정 요약 정보 반환"""
        try:
            config = self.get_config()
            whitelist_rules = self.get_whitelist()
            blacklist_patterns = self.get_blacklist()
            
            return {
                'config_version': config.get('version', 'unknown'),
                'config_directory': str(self.config_dir),
                'whitelist_rules_count': len(whitelist_rules),
                'blacklist_patterns_count': len(blacklist_patterns),
                'auto_backup_enabled': config['general']['auto_backup'],
                'default_strategy': config['processing']['default_strategy'],
                'security_level': 'high' if not config['security']['log_sensitive_data'] else 'standard'
            }
            
        except Exception as e:
            self.logger.error(f"설정 요약 생성 실패: {e}")
            return {'error': str(e)}


# 전역 설정 관리자 인스턴스
_config_manager = None

def get_config_manager() -> ConfigManager:
    """싱글톤 설정 관리자 반환"""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


if __name__ == "__main__":
    # 테스트용 코드
    try:
        config_mgr = ConfigManager()
        
        # 설정 로드 테스트
        config = config_mgr.get_config()
        print("메인 설정 로드 성공")
        
        whitelist = config_mgr.get_whitelist()
        print(f"화이트리스트 로드 성공: {len(whitelist)}개 규칙")
        
        blacklist = config_mgr.get_blacklist()
        print(f"블랙리스트 로드 성공: {len(blacklist)}개 패턴")
        
        # 유효성 검사
        validation_results = config_mgr.validate_all_configs()
        print(f"설정 유효성 검사: {validation_results}")
        
        # 요약 정보
        summary = config_mgr.get_config_summary()
        print(f"설정 요약: {summary}")
        
        print("\n보안 강화된 설정 관리 시스템 초기화 완료!")
        
    except Exception as e:
        print(f"초기화 실패: {e}")