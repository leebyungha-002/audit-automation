"""
감사 녹음 파일 → 텍스트 변환기 (faster-whisper)

사용법:
    python transcribe.py                     # input/ 폴더의 미처리 음성파일 전체 변환
    python transcribe.py --model medium      # 모델 변경 (tiny/base/small/medium/large-v3/large-v3-turbo)
    python transcribe.py --force             # 이미 변환된 파일도 다시 변환
    python transcribe.py --file 회의.m4a      # 특정 파일만 변환

출력 (output/ 폴더):
    <파일명>.txt  — [시:분:초] 타임스탬프가 붙은 대화록
    <파일명>.srt  — 자막 형식 (녹음 재생하며 확인할 때 사용)
"""
import argparse
import os
import sys
import time

# 런처/콘솔(cp949)에서 특수문자 출력 시 UnicodeEncodeError 방지
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)  # 런처에서 진행률 즉시 표시
    except Exception:
        pass

# 모델 다운로드 시 나오는 무해한 경고 숨김 (Windows 심볼릭링크 미지원, HF 토큰 미설정)
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
MODEL_DIR = os.path.join(BASE_DIR, "models")

AUDIO_EXTS = {".mp3", ".m4a", ".wav", ".ogg", ".flac", ".aac", ".wma", ".webm", ".mp4", ".amr", ".3gp"}

# 감사 용어 인식률을 높이기 위한 힌트 (whisper initial_prompt)
AUDIT_PROMPT = (
    "회계감사 인터뷰 녹음입니다. 매출채권, 대손충당금, 재고자산, 유형자산, 감가상각, "
    "리스부채, 사용권자산, 퇴직급여충당부채, 연차충당부채, 차입금, 이자비용, 분개장, "
    "계정별원장, 결산, 전표, 거래처, 세금계산서, 내부통제, 재무제표, 감사조서."
)


def fmt_ts(sec, srt=False):
    h, rem = divmod(int(sec), 3600)
    m, s = divmod(rem, 60)
    if srt:
        ms = int(round((sec - int(sec)) * 1000))
        return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"
    return f"{h:02d}:{m:02d}:{s:02d}"


def patch_pyav():
    """PyAV 버전에 따라 av.open()이 metadata_errors 인자를 거부하는 문제 우회
    (미니PC에서 'open() got an unexpected keyword argument metadata_errors' 발생)"""
    import av

    orig_open = av.open

    def safe_open(*args, **kwargs):
        try:
            return orig_open(*args, **kwargs)
        except TypeError as e:
            if "metadata_errors" not in str(e):
                raise
            kwargs.pop("metadata_errors", None)
            return orig_open(*args, **kwargs)

    av.open = safe_open


def load_model(model_name):
    from faster_whisper import WhisperModel
    try:
        from huggingface_hub.utils import logging as hf_logging
        hf_logging.set_verbosity_error()
    except Exception:
        pass

    device, compute_type = "cpu", "int8"
    try:
        import ctranslate2
        if ctranslate2.get_cuda_device_count() > 0:
            device, compute_type = "cuda", "float16"
    except Exception:
        pass

    print(f"[모델 로드] {model_name} ({device}/{compute_type}) — 최초 실행 시 모델 다운로드로 시간이 걸립니다")
    os.makedirs(MODEL_DIR, exist_ok=True)
    return WhisperModel(
        model_name,
        device=device,
        compute_type=compute_type,
        cpu_threads=os.cpu_count() or 4,
        download_root=MODEL_DIR,
    )


def transcribe_file(model, src, language):
    name = os.path.splitext(os.path.basename(src))[0]
    txt_path = os.path.join(OUTPUT_DIR, f"{name}.txt")
    srt_path = os.path.join(OUTPUT_DIR, f"{name}.srt")

    size_mb = os.path.getsize(src) / 1024 / 1024
    print(f"\n[변환 시작] {os.path.basename(src)} ({size_mb:.1f}MB)")
    started = time.time()

    segments, info = model.transcribe(
        src,
        language=language,
        beam_size=5,
        vad_filter=True,                      # 무음 구간 제거 → 속도↑, 환각 문장↓
        vad_parameters={"min_silence_duration_ms": 500},
        condition_on_previous_text=False,     # 같은 문장 반복 생성 방지
        initial_prompt=AUDIT_PROMPT,
    )
    duration = info.duration or 0
    print(f"  녹음 길이: {fmt_ts(duration)}  / 언어: {info.language} ({info.language_probability:.0%})")

    txt_lines, srt_lines = [], []
    last_pct = -10
    for i, seg in enumerate(segments, 1):
        text = seg.text.strip()
        if not text:
            continue
        txt_lines.append(f"[{fmt_ts(seg.start)}] {text}")
        srt_lines.append(f"{i}\n{fmt_ts(seg.start, True)} --> {fmt_ts(seg.end, True)}\n{text}\n")
        if duration:
            pct = int(seg.end / duration * 100)
            if pct >= last_pct + 10:
                last_pct = pct - pct % 10
                elapsed = time.time() - started
                print(f"  진행 {pct:3d}%  ({fmt_ts(seg.end)} / {fmt_ts(duration)}, 경과 {fmt_ts(elapsed)})")

    header = [
        f"원본 파일: {os.path.basename(src)}",
        f"녹음 길이: {fmt_ts(duration)}",
        f"변환 일시: {time.strftime('%Y-%m-%d %H:%M')}",
        "",
    ]
    with open(txt_path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(header + txt_lines) + "\n")
    with open(srt_path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(srt_lines))

    elapsed = time.time() - started
    print(f"[완료] {len(txt_lines)}개 문장, 소요 {fmt_ts(elapsed)} → output/{name}.txt")


def main():
    ap = argparse.ArgumentParser(description="감사 녹음 파일 텍스트 변환 (faster-whisper)")
    ap.add_argument("--model", default="large-v3-turbo", help="whisper 모델 (기본: large-v3-turbo)")
    ap.add_argument("--lang", default="ko", help="언어 코드 (기본: ko, 자동감지는 auto)")
    ap.add_argument("--file", help="input 폴더 내 특정 파일만 변환")
    ap.add_argument("--force", action="store_true", help="이미 변환된 파일도 다시 변환")
    args = ap.parse_args()

    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if args.file:
        files = [os.path.join(INPUT_DIR, args.file)]
        if not os.path.exists(files[0]):
            sys.exit(f"[오류] 파일이 없습니다: {files[0]}")
    else:
        files = sorted(
            os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR)
            if os.path.splitext(f)[1].lower() in AUDIO_EXTS
        )

    if not args.force:
        todo = []
        for f in files:
            name = os.path.splitext(os.path.basename(f))[0]
            if os.path.exists(os.path.join(OUTPUT_DIR, f"{name}.txt")):
                print(f"[건너뜀] 이미 변환됨: {os.path.basename(f)} (다시 하려면 --force)")
            else:
                todo.append(f)
        files = todo

    if not files:
        print(f"변환할 음성 파일이 없습니다. input 폴더에 파일을 넣어주세요:\n  {INPUT_DIR}")
        return

    patch_pyav()
    model = load_model(args.model)
    language = None if args.lang == "auto" else args.lang
    for f in files:
        try:
            transcribe_file(model, f, language)
        except Exception as e:
            print(f"[실패] {os.path.basename(f)}: {e}")


if __name__ == "__main__":
    main()
