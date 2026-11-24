"""Retrieve current approved files, then build a new local preview. Does not publish."""
import argparse
from pathlib import Path
from tempfile import TemporaryDirectory
import build
import snapshot

if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--output',type=Path,default=Path('site-output'))
    ap.add_argument('--report',type=Path,default=Path('review-report.json'))
    args=ap.parse_args()
    output=args.output.resolve(); report=args.report.resolve()
    if output.exists() or report.exists(): ap.error('Choose new paths; this command will not replace existing files.')
    output.parent.mkdir(parents=True,exist_ok=True)
    report.parent.mkdir(parents=True,exist_ok=True)
    with TemporaryDirectory(prefix='fedoras-stage-',dir=output.parent) as td:
        stage=Path(td)
        snapshot.snapshot(stage/'snapshot')
        build.build(stage/'snapshot',stage/'public',stage/'review-report.json',build.ROOT/'approved-file-list.json')
        (stage/'review-report.json').replace(report)
        (stage/'public').rename(output)
    print(f'Local site: {output}\nReview report (do not deploy): {report}')
