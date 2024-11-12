# VCP proces

```mermaid
sequenceDiagram
    participant Lid
    participant VCP
    participant Bestuur

    activate Lid
    Lid ->> VCP : Melding van ongepast gedrag
    activate VCP
    VCP -->> Bestuur : Er is een melding
    activate Bestuur

    Bestuur -->> VCP : Ondernomen actie
    deactivate Bestuur
    VCP ->> Lid : Feedback
    deactivate VCP
    deactivate Lid
```
