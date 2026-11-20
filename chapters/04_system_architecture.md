# CHAPTER 4: SYSTEM ARCHITECTURE & DESIGN

## 4.1 Tiered Component Structure
The system is partitioned into three decoupled functional tiers:
1. **Presentation Tier:** Responsive Single Page Application providing interactive spatial mapping and parameter simulation.
2. **Microservice Application Tier:** High-throughput REST API gateway ensuring sub-20ms inference latency.
3. **Soft Computing Analytical Core:** Specialized numerical computation libraries managing ML inferences and Pareto evaluations.
