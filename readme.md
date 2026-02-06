# Classic ASP Professional Benchmark Suite

A high-precision micro-benchmarking tool for VBScript (Classic ASP). This suite allows developers to compare the performance of different code blocks, iteration methods, and logic patterns with professional-grade accuracy.

## 🤖 Attribution
**Note:** This project was co-authored by a Human Developer and an Artificial Intelligence. 
- **Human Supervision:** Architecture design, feature requirements, logic validation, and VBScript scope troubleshooting.
- **AI Implementation:** Code generation, syntax hardening, and documentation formatting.

## ✨ Key Features
- **Hardened Execution:** Uses `bm_` variable prefixes to prevent namespace collisions during `ExecuteGlobal`.
- **Automatic Warmup:** CPU/P-Code stabilization before timing begins.
- **Multi-Round Sampling:** Run tests multiple times (Rounds) and automatically select the "Best Time" to eliminate server noise.
- **Flexible Streams:** Output raw data to `FILE`, `CASHE`, or `STR`, and human-readable tables to `SCREEN`.
- **Article Ready:** Built-in `getHTMLTable()` with CSS classes for seamless integration into technical blogs or documentation.

## 🚀 Quick Start

```asp
<!--#include file="Benchmark.class.asp"-->
<%
Dim Bench : Set Bench = New Benchmark_Class

Bench.Name = "Loop_Test"
Bench.Rounds = 3
Bench.SetMode "COUNT", 100000

' Setup environment
Bench.InitGlobal = "Dim arr(99), i_t, x_t : For i_t = 0 To 99 : arr(i_t) = i_t : Next"
Bench.InitBlock  = "i_t = 0 : x_t = Empty"

' Add tests
Bench.AddCodeBlock "For_Each", "For Each item In arr : x_t = item : Next"
Bench.AddCodeBlock "For_Next", "For i_t = 0 To 99 : x_t = arr(i_t) : Next"

Bench.Run
Response.Write Bench.getHTMLTable()
