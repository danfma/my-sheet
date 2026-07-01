```

BenchmarkDotNet v0.15.8, macOS Tahoe 26.5.1 (25F80) [Darwin 25.5.0]
Apple M1 Pro, 1 CPU, 10 logical and 10 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.9, 10.0.926.27113), Arm64 RyuJIT armv8.0-a
  ShortRun : .NET 10.0.9 (10.0.9, 10.0.926.27113), Arm64 RyuJIT armv8.0-a

Job=ShortRun  IterationCount=3  LaunchCount=1  
WarmupCount=3  Categories=CumChain  

```
| Method             | Depth | Mean      | Error      | StdDev    | Ratio | RatioSD | Gen0    | Allocated | Alloc Ratio |
|------------------- |------ |----------:|-----------:|----------:|------:|--------:|--------:|----------:|------------:|
| **CumChain_Object**    | **1000**  | **13.309 μs** |  **4.9264 μs** | **0.2700 μs** |  **1.00** |    **0.02** |  **7.6447** |   **47976 B** |        **1.00** |
| CumChain_BoxCache  | 1000  | 13.338 μs | 10.3582 μs | 0.5678 μs |  1.00 |    0.04 |  6.5918 |   41352 B |        0.86 |
| CumChain_CellValue | 1000  |  5.853 μs |  1.0298 μs | 0.0564 μs |  0.44 |    0.01 |       - |         - |        0.00 |
|                    |       |           |            |           |       |         |         |           |             |
| **CumChain_Object**    | **3000**  | **44.429 μs** |  **2.6303 μs** | **0.1442 μs** |  **1.00** |    **0.00** | **22.9492** |  **143976 B** |        **1.00** |
| CumChain_BoxCache  | 3000  | 43.743 μs |  3.0926 μs | 0.1695 μs |  0.98 |    0.00 | 21.8506 |  137352 B |        0.95 |
| CumChain_CellValue | 3000  | 18.447 μs |  0.1148 μs | 0.0063 μs |  0.42 |    0.00 |       - |         - |        0.00 |
