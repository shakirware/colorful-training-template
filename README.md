# Colorful Training Template

This tool generates structured workout templates using YAML input data. It calculates one-rep max percentages using `.xlsx` files sourced from StrengthLevel.

## Usage

1. Input workout data in YAML format, specifying:
   - **Exercise names**
   - **Sets and reps**
   - **Percentage of 1RM** (if applicable)
   - **Weights** (automatically computed based on StrengthLevel data)

2. The tool will process the input and generate a structured training program.

3. Outputs a visually formatted plan that is free for anyone to use.

---

## Example YAML Input

```yaml
- week_1:
    # Transitional Week – Ease into the volume block with 2 sessions.
    - weekday: Friday
      session: Lower Body (Volume Intro)
      exercises:
        - name: Squat
          sets:
            - notes: Standard – focus on technique and controlled pace.
              percentage_1rm: 60
              reps: "2x5"
            - notes: Paused variation – emphasize control.
              percentage_1rm: 55
              reps: "2x2"
            - notes: Explosive reps – maintain speed without sacrificing form.
              percentage_1rm: 55
              reps: "2x3"
        - name: Leg Press
          sets:
            - notes: Focus on full range of motion.
              reps: "2x10"
    - weekday: Saturday
      session: Upper Body (Volume Intro)
      exercises:
        - name: Weighted Muscle-Ups
          sets:
            - notes: Keep it moderate; focus on smooth movement.
              percentage_1rm: 60
              reps: "2x5"
        - name: Weighted Pull-Ups
          sets:
            - notes: Standard execution – controlled and quality reps.
              percentage_1rm: 65
              reps: "3x5"
              weight: "47.5kg"
```
---


## Screenshot

![image](https://github.com/user-attachments/assets/72365b5c-1801-45e5-91f0-9af3ade5b55b)

