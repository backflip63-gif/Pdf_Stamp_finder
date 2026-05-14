from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable, List

import fitz

from .analyzer import PageAnalyzer
from .models import BatchJobConfig, FileProcessResult, PlacementResult
from .placer import StampPlacer
from .stamp_utils import get_stamp_page_size, place_stamp_pdf
from .utils import build_output_path, ensure_dir, mm_to_pt


class BatchProcessor:
    def __init__(self, config: BatchJobConfig):
        self.config = config
        self.analyzer = PageAnalyzer(
            render_dpi=config.settings.render_dpi,
            whiteness_threshold=config.settings.whiteness_threshold,
            dilation_px=config.settings.dilation_px,
        )
        self.placer = StampPlacer(config.settings)

    def iter_input_pdfs(self) -> Iterable[Path]:
        for path in sorted(self.config.input_dir.glob("*.pdf")):
            if path.is_file():
                yield path

    def process_all(self, progress_callback: Callable[[int, int, Path, FileProcessResult], None] | None = None) -> List[FileProcessResult]:
        ensure_dir(self.config.output_dir)
        pdfs = list(self.iter_input_pdfs())
        total = len(pdfs)
        results: List[FileProcessResult] = []
        for index, pdf_path in enumerate(pdfs, start=1):
            file_result = self.process_file(pdf_path)
            results.append(file_result)
            if progress_callback:
                progress_callback(index, total, pdf_path, file_result)
        return results

    def process_file(self, pdf_path: Path) -> FileProcessResult:
        output_file = build_output_path(pdf_path, self.config.output_dir, self.config.output_suffix)
        result = FileProcessResult(input_file=pdf_path, output_file=output_file)

        try:
            stamp_doc = fitz.open(self.config.stamp_pdf)
            stamp_src_w, stamp_src_h = get_stamp_page_size(self.config.stamp_pdf)
            stamp_source = stamp_doc[0].get_pixmap(dpi=300, alpha=True)

            # Gewünschte Zielgröße ist konfigurierbar. Wenn 0, nimm Originalgröße.
            target_w = mm_to_pt(self.config.settings.stamp_width_mm) if self.config.settings.stamp_width_mm > 0 else stamp_src_w
            target_h = mm_to_pt(self.config.settings.stamp_height_mm) if self.config.settings.stamp_height_mm > 0 else stamp_src_h

            doc = fitz.open(pdf_path)
            try:
                pages_to_process = self._page_indices(len(doc))
                for page_index in pages_to_process:
                    page = doc[page_index]
                    analysis = self.analyzer.analyze(page)
                    cand = self.placer.find_position(
                        page,
                        analysis.occupancy_mask,
                        analysis.zoom,
                        target_w,
                        target_h,
                    )
                    if cand is None:
                        result.page_results.append(
                            PlacementResult(
                                page_index=page_index,
                                rect=None,
                                scale=1.0,
                                occupancy_ratio=1.0,
                                status="no_position",
                                note="Keine ausreichend freie Fläche gefunden.",
                            )
                        )
                        continue

                    place_stamp_pdf(page, stamp_source, cand.rect)
                    result.page_results.append(
                        PlacementResult(
                            page_index=page_index,
                            rect=(cand.rect.x0, cand.rect.y0, cand.rect.x1, cand.rect.y1),
                            scale=cand.scale,
                            occupancy_ratio=cand.occupancy_ratio,
                            status="placed",
                            note="",
                        )
                    )

                doc.save(output_file, garbage=4, deflate=True)
                result.success = True
            finally:
                doc.close()
                stamp_doc.close()
        except Exception as exc:
            result.success = False
            result.error = str(exc)

        return result

    def _page_indices(self, page_count: int) -> List[int]:
        mode = self.config.settings.process_mode
        if page_count <= 0:
            return []
        if mode == "first":
            return [0]
        if mode == "last":
            return [page_count - 1]
        return list(range(page_count))
