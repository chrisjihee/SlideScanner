from pptx import Presentation

from chrisbase.data import *
from chrisbase.io import *
from chrisbase.util import *

logger = logging.getLogger(__name__)
args = CommonArguments(
    env=ProjectEnv(
        project="SlideScanner",
        job_name=Path(__file__).stem,
        msg_level=logging.INFO,
        msg_format=LoggingFormat.PRINT_00,
    )
)
args.info_args()


def scan_pptx(file_path):
    file_cont = list()
    for slide in Presentation(file_path).slides:
        slide_cont = []
        for shape in slide.shapes:
            shape_cont = {"name": shape.name, "text": str(shape.text).strip().replace("\r", "").replace("\n", '<BR>')}
            slide_cont.append(shape_cont)
        file_cont.append(slide_cont)
    return file_cont


with JobTimer("WritingAnalysis", rt=1, rb=1, rw=80, rc='=', verbose=1):
    # input_files = sorted(Path("/Users/chris/Seafile/love/찬양 PPT").glob("*.pptx"))[:1]
    input_files = sorted(Path("/Users/chris/Seafile/temp/찬양 PPT").glob("*.pptx"))

    for file in input_files:
        print(file)
        contents = scan_pptx(file)
        print(json.dumps(contents, indent=4, ensure_ascii=False))
