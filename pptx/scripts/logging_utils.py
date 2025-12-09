#!/usr/bin/env python3
"""Logging utilities for PPTX generation.

Logs are written to:
- {working_dir}/powerpoint/processing/pptx_generation.log (new structure, recommended)
- {working_dir}/processing/pptx_generation.log (legacy)
"""

import logging
import os
from pathlib import Path
from typing import Optional


class PPTXLogger:
    """Logger for PPTX generation with automatic log file placement."""

    _instance: Optional[logging.Logger] = None
    _log_file_path: Optional[Path] = None

    @classmethod
    def _detect_working_dir(cls) -> Path:
        """Auto-detect working directory.

        Search order:
        1. Current working directory if it contains 'powerpoint/processing/'
        2. Current working directory if it contains 'processing/' or is named 'presentation'
        3. Parent directory if it's named 'presentation'
        4. ./presentation/ if it exists
        5. Current working directory (fallback)
        """
        cwd = Path.cwd()

        # Check if current dir has powerpoint/processing (new structure)
        if (cwd / 'powerpoint' / 'processing').exists():
            return cwd

        # Check if current dir is presentation or has processing (legacy)
        if cwd.name == 'presentation' or (cwd / 'processing').exists():
            return cwd

        # Check if parent is presentation
        if cwd.parent.name == 'presentation':
            return cwd.parent

        # Check for ./presentation/
        if (cwd / 'presentation').exists() and (cwd / 'presentation').is_dir():
            return cwd / 'presentation'

        # Fallback to cwd
        return cwd

    @classmethod
    def setup(cls, working_dir: Optional[str] = None):
        """Setup logger with file handler in processing directory.

        Args:
            working_dir: Working directory (e.g., project root or presentation/).
                        If None, auto-detects or uses current working directory.
        """
        if cls._instance is not None:
            return cls._instance

        # Determine log file location
        if working_dir:
            base_dir = Path(working_dir)
        else:
            base_dir = cls._detect_working_dir()

        # Prefer powerpoint/processing/ (new structure), fallback to processing/ (legacy)
        if (base_dir / 'powerpoint' / 'processing').exists() or (base_dir / 'powerpoint').exists():
            processing_dir = base_dir / "powerpoint" / "processing"
        elif (base_dir / 'processing').exists():
            processing_dir = base_dir / "processing"
        else:
            # Default to new structure
            processing_dir = base_dir / "powerpoint" / "processing"

        processing_dir.mkdir(parents=True, exist_ok=True)

        cls._log_file_path = processing_dir / "pptx_generation.log"

        # Create logger
        logger = logging.getLogger('pptx_generation')
        logger.setLevel(logging.DEBUG)

        # Remove existing handlers
        logger.handlers = []

        # File handler
        file_handler = logging.FileHandler(cls._log_file_path, mode='a', encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # Console handler (warnings and errors only)
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)

        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        cls._instance = logger

        logger.info("=" * 60)
        logger.info("PPTX Generation Started")
        logger.info(f"Log file: {cls._log_file_path}")
        logger.info("=" * 60)

        return logger

    @classmethod
    def get_logger(cls) -> logging.Logger:
        """Get or create logger instance."""
        if cls._instance is None:
            return cls.setup()
        return cls._instance

    @classmethod
    def get_log_path(cls) -> Optional[Path]:
        """Get path to log file."""
        return cls._log_file_path


def get_logger() -> logging.Logger:
    """Convenience function to get logger."""
    return PPTXLogger.get_logger()


if __name__ == '__main__':
    # Test logging
    logger = PPTXLogger.setup(working_dir='.')
    logger.debug("Debug message")
    logger.info("Info message")
    logger.warning("Warning message")
    logger.error("Error message")
    print(f"Log file created at: {PPTXLogger.get_log_path()}")
