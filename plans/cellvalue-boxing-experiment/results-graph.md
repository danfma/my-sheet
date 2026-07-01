```

BenchmarkDotNet v0.15.8, macOS Tahoe 26.5.1 (25F80) [Darwin 25.5.0]
Apple M1 Pro, 1 CPU, 10 logical and 10 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.9, 10.0.926.27113), Arm64 RyuJIT armv8.0-a
  ShortRun : .NET 10.0.9 (10.0.9, 10.0.926.27113), Arm64 RyuJIT armv8.0-a

Job=ShortRun  IterationCount=3  LaunchCount=1  
WarmupCount=3  Categories=Graph  

```
| Method         | Cells  | Mean       | Error       | StdDev    | Ratio | RatioSD | Gen0     | Gen1     | Allocated | Alloc Ratio |
|--------------- |------- |-----------:|------------:|----------:|------:|--------:|---------:|---------:|----------:|------------:|
| **Graph_ObjCache** | **10000**  |   **477.1 μs** |    **77.04 μs** |   **4.22 μs** |  **1.00** |    **0.01** |  **38.0859** |   **9.2773** |  **239904 B** |        **1.00** |
| Graph_CvCache  | 10000  |   456.0 μs |   125.17 μs |   6.86 μs |  0.96 |    0.01 |        - |        - |         - |        0.00 |
|                |        |            |             |           |       |         |          |          |           |             |
| **Graph_ObjCache** | **100000** | **5,642.0 μs** | **1,078.85 μs** |  **59.14 μs** |  **1.00** |    **0.01** | **375.0000** | **171.8750** | **2399904 B** |        **1.00** |
| Graph_CvCache  | 100000 | 4,948.7 μs | 2,271.90 μs | 124.53 μs |  0.88 |    0.02 |        - |        - |         - |        0.00 |
