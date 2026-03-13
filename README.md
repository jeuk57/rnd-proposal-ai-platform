# rnd-proposal-ai-platform

Public portfolio extraction of the PPT generation pipeline from a larger R&D proposal automation project.

## Structure
- `src/ppt_maker/main_ppt.py`: LangGraph-based PPT pipeline entrypoint
- `src/ppt_maker/nodes/`: section split, generation, render, and postprocess nodes
- `src/ppt_maker/background/`: background assets used in the postprocess layer
- `src/utils/`: minimal parsing and DB lookup helpers required by the pipeline

## Run
Use the repository root as the working directory.

```bash
python -m src.ppt_maker.main_ppt --help
```
