#!/bin/bash

# Create output directory if it doesn't exist
mkdir -p output_tests

# Array of input files
files=(
    "samples/JS_WORK-2207433-v13A-_효성비나제일차_지분매매계약_[보통주 주요 조건].DOCX"
    "samples/JS_WORK-2207483-v19A-_효성비나제일차_정산계약_[주가수익스왑계약(PRS)].DOCX"
    "samples/JS_WORK-2214797-v10A-_효성비나제일차__ABL대출약정_.docx"
    "samples/JS_WORK-2250690-v2A-_KB스타리츠_대출약정서.DOCX"
    "samples/포항장성동_영검보.docx"
    "samples/2. 투자계약서_DB Carlyle 인프라 일반사모투자신탁제1호_vF.docx"
    "samples/[날인예정본]효성비나제이차_유동화증권 인수약정_[유동화증권].DOCX"
    "samples/[날인예정본]효성비나제이차_사모사채 인수확약 및 자금보충 합의서[사모사채인수확약].docx"
)

# Build project first
echo "Building project..."
cargo build --release

# Create output directory
mkdir -p output_tests/samples || { echo "Build failed"; exit 1; }

echo "Starting tests..."
echo "----------------------------------------"

for input_file in "${files[@]}"; do
    if [ ! -f "$input_file" ]; then
        echo "⚠️  File not found: $input_file"
        continue
    fi
    
    output_file="output_tests/${input_file%.*}.md"
    echo "Processing: $input_file"
    
    if cargo run --release -- "$input_file" "$output_file"; then
        echo "✅ Success"
    else
        echo "❌ Failed"
    fi
    echo "----------------------------------------"
done
