# Rapport de standardisation — Numérotation et niveaux de titres

## 1. Objet du rapport

Ce rapport établit le standard de numérotation et de hiérarchie Markdown à appliquer au manuscrit **Guide — Du savoir à l’action**.

Il sert de référence à toute IA chargée de corriger, restructurer ou prolonger le livre.

Objectif principal :

> Uniformiser la structure du livre avec une numérotation stable à trois niveaux maximum, de type `3.3.3`, et aligner les niveaux Markdown sur la hiérarchie réelle du contenu.

---

## 2. Diagnostic général

Le manuscrit contient une logique pédagogique cohérente, mais la structure Markdown est instable.

Problèmes observés :

1. Trop de titres utilisent `#`, alors que `#` devrait être réservé aux titres de chapitre.
2. Certaines sections internes sont numérotées comme des chapitres.
3. Certains chapitres utilisent la numérotation locale `1.`, `2.`, `3.` au lieu de `7.1`, `7.2`, `7.3`.
4. Certains sous-titres sont au mauvais niveau Markdown.
5. Des blocs comme `Fonction UCKK dominante` ou `Transformation du chapitre` sont parfois traités comme des titres, alors qu’ils devraient être des métadonnées de chapitre.
6. Des exemples, exercices ou tableaux sont parfois promus à un niveau de titre trop élevé.
7. Des titres de niveau 4 existent alors que le standard cible doit rester à trois niveaux maximum.

Conséquence :

> La table des matières devient confuse, les chapitres semblent se fragmenter, et l’IA risque de prolonger les incohérences au lieu de les corriger.

---

## 3. Standard retenu

Le standard retenu est une hiérarchie à trois niveaux numérotés.

### Niveau 1 — Chapitre

Format Markdown :

```md
# 03. Inventorier les ressources
```

Règle :

- Un seul `#` par chapitre.
- Le numéro de chapitre est toujours sur deux chiffres de `01` à `18`.
- Ne jamais utiliser `#` pour une section interne.
- Ne jamais écrire `# 3.1`, `# 4.2`, `# 10.12`, etc.

Usage :

```md
# 01. Introduction générale
# 02. Commencer par ce qui existe déjà
# 03. Inventorier les ressources
# 04. Écouter avec attention
```

---

### Niveau 2 — Section principale du chapitre

Format Markdown :

```md
## 3.1 Situer — Avant d’agir, savoir ce qui existe
```

Règle :

- Le niveau 2 correspond aux grandes sections internes.
- La numérotation commence toujours par le numéro du chapitre.
- Le format est `chapitre.section`.
- Le numéro de section suit l’ordre réel du chapitre.
- Les sections standards du guide doivent rester cohérentes d’un chapitre à l’autre.

Exemples corrects :

```md
## 3.0 Objectifs du chapitre
## 3.1 Situer — Avant d’agir, savoir ce qui existe
## 3.2 Distinguer — Ressource, besoin, solution
## 3.3 Transformer — De la liste à la carte
## 3.4 Produire — Carte des ressources
## 3.5 Retenir — Ce qu’une carte permet et ne permet pas
## 3.6 Référencer — Références mobilisées
## 3.7 Livrable du chapitre
## 3.8 Transition vers le chapitre 4
```

---

### Niveau 3 — Sous-section

Format Markdown :

```md
### 3.3.3 Le cas COVID comme exercice de lucidité
```

Règle :

- Le niveau 3 correspond à une subdivision directe d’une section.
- Le format est `chapitre.section.sous-section`.
- C’est le dernier niveau numéroté autorisé.
- La numérotation à trois niveaux est le modèle cible.

Exemples corrects :

```md
### 3.1.1 Idée centrale
### 3.1.2 Place du chapitre dans le guide
### 3.1.3 Application au Grand Dossier COVID
### 3.3.1 Opération centrale
### 3.3.2 Pourquoi cette transformation est nécessaire
### 3.3.3 Le cas COVID comme exercice de lucidité
```

---

## 4. Interdiction du niveau 4 numéroté

Le standard cible ne doit pas dépasser trois niveaux numérotés.

À éviter :

```md
#### 3.3.3.1 Exemple
#### Étape 1 — Nommer l’action prévue
```

À utiliser plutôt :

```md
**Exemple — Écoles et jeunes**

**Étape 1 — Nommer l’action prévue**
```

Règle :

- Les exemples, étapes, notes pédagogiques, tableaux, consignes, variantes et points de vigilance ne doivent pas créer un quatrième niveau de hiérarchie.
- Ils doivent être présentés en gras, en liste ou en encadré, mais pas en heading Markdown.

---

## 5. Métadonnées de chapitre

Les éléments suivants ne doivent pas être des titres Markdown :

- Fonction UCKK dominante
- Transformation du chapitre
- Objectif pédagogique court
- Formule directrice
- Intention du chapitre

Format recommandé :

```md
# 03. Inventorier les ressources

**Fonction UCKK dominante :** Situer → Produire

**Transformation du chapitre :**  
situation vague  
→ ressources repérées  
→ ressources classées  
→ porteurs identifiés  
→ accès clarifiés  
→ limites reconnues  
→ première carte d’action possible
```

Raison :

> Ces éléments cadrent le chapitre, mais ne doivent pas apparaître comme sections principales dans la table des matières.

Exception possible :

- `## 3.0 Objectifs du chapitre` peut rester une section officielle.
- Les objectifs sont une vraie section pédagogique, contrairement aux métadonnées courtes.

---

## 6. Structure canonique d’un chapitre UCKK

Chaque chapitre doit suivre autant que possible cette structure.

```md
# XX. Titre du chapitre

**Fonction UCKK dominante :** [fonction]

**Transformation du chapitre :**  
[état initial]  
→ [étape intermédiaire]  
→ [état visé]

## XX.0 Objectifs du chapitre

## XX.1 Situer — [titre contextualisé]

## XX.2 Distinguer — [titre contextualisé]

## XX.3 Transformer — [titre contextualisé]

## XX.4 Produire — [livrable ou exercice principal]

## XX.5 Retenir — [points clés]

## XX.6 Référencer — [références mobilisées]

## XX.7 Livrable du chapitre

## XX.8 Transition vers le chapitre suivant
```

Cette structure peut être adaptée, mais l’ordre UCKK doit rester lisible.

---

## 7. Structure canonique d’un chapitre de synthèse

Les chapitres de synthèse peuvent avoir une structure légèrement différente, mais doivent respecter la même numérotation.

Exemple :

```md
# 11. Synthèse de la partie 2

**Fonction UCKK dominante :** Évaluer → Mémoriser

## 11.0 Objectifs de la synthèse

## 11.1 Situer — De la compréhension à l’orientation

## 11.2 Distinguer — Ne pas confondre option, décision et action

## 11.3 Transformer — De l’inventaire des solutions au choix justifié

## 11.4 Produire — Note de décision

## 11.5 Retenir — Phrase clé et erreurs fréquentes

## 11.6 Référencer — Références mobilisées

## 11.7 Transition vers le chapitre 12
```

À corriger :

```md
# 1. Situer — De la compréhension à l’orientation
```

Devient :

```md
## 11.1 Situer — De la compréhension à l’orientation
```

---

## 8. Règles de correction automatique pour l’IA

Toute IA qui corrige le livre doit appliquer les règles suivantes.

### Règle 1 — Ne jamais multiplier les H1

Un chapitre ne doit avoir qu’un seul titre de niveau `#`.

Si un titre interne est actuellement en `#`, le rétrograder.

Exemple :

```md
# 3.1 Situer — Avant d’agir, savoir ce qui existe
```

Devient :

```md
## 3.1 Situer — Avant d’agir, savoir ce qui existe
```

---

### Règle 2 — Toujours rattacher la section au numéro du chapitre

Si le chapitre est 07, toutes les sections internes doivent commencer par `7.`.

À corriger :

```md
# 1. Situer — Le réflexe de solution rapide
# 2. Distinguer — Ce qu’il ne faut pas confondre
```

Devient :

```md
## 7.1 Situer — Le réflexe de solution rapide
## 7.2 Distinguer — Ce qu’il ne faut pas confondre
```

---

### Règle 3 — Décaler les niveaux Markdown selon la hiérarchie réelle

Si une section est `3.1`, elle doit être en H2.

Si une sous-section est `3.1.1`, elle doit être en H3.

Exemple :

```md
# 3.1 Situer — Avant d’agir, savoir ce qui existe
## 3.1.1 Idée centrale
```

Devient :

```md
## 3.1 Situer — Avant d’agir, savoir ce qui existe
### 3.1.1 Idée centrale
```

---

### Règle 4 — Supprimer les H4

Tout titre `####` doit être converti en gras ou en élément de liste.

Exemple :

```md
#### Exemple — Écoles et jeunes
```

Devient :

```md
**Exemple — Écoles et jeunes**
```

---

### Règle 5 — Les listes internes ne deviennent pas des titres numérotés

Dans une section comme `3.2.3 Les dix types de ressources`, les dix types ne doivent pas devenir des sous-sections de même niveau que `3.2.3`.

À éviter :

```md
### 1. Ressources communautaires
### 2. Ressources individuelles
```

À préférer :

```md
**Type 1 — Ressources communautaires**

**Type 2 — Ressources individuelles**
```

ou :

```md
1. **Ressources communautaires** — ...
2. **Ressources individuelles** — ...
```

---

### Règle 6 — Les transitions doivent être numérotées dans le chapitre

À corriger :

```md
# Transition vers le chapitre 12
```

Devient :

```md
## 11.7 Transition vers le chapitre 12
```

---

### Règle 7 — Les livrables doivent être numérotés dans le chapitre

À corriger :

```md
## Livrable du chapitre
```

Devient :

```md
## 16.6 Livrable du chapitre
```

Le numéro exact dépend de la place réelle dans le chapitre.

---

## 9. Modèle corrigé complet

Voici un modèle de chapitre conforme au standard.

```md
# 03. Inventorier les ressources

**Fonction UCKK dominante :** Situer → Produire

**Transformation du chapitre :**  
situation vague  
→ ressources repérées  
→ ressources classées  
→ porteurs identifiés  
→ accès clarifiés  
→ limites reconnues  
→ première carte d’action possible

## 3.0 Objectifs du chapitre

À la fin du chapitre, l’étudiant devrait être capable de...

## 3.1 Situer — Avant d’agir, savoir ce qui existe

### 3.1.1 Idée centrale

### 3.1.2 Place du chapitre dans le guide

### 3.1.3 Application au Grand Dossier COVID

## 3.2 Distinguer — Ressource, besoin, solution

### 3.2.1 Distinctions importantes

### 3.2.2 Définition de ressource

### 3.2.3 Les dix types de ressources

**Type 1 — Ressources communautaires**

**Type 2 — Ressources individuelles**

**Type 3 — Ressources relationnelles**

## 3.3 Transformer — De la liste à la carte

### 3.3.1 Opération centrale

### 3.3.2 Pourquoi cette transformation est nécessaire

### 3.3.3 Le cas COVID comme exercice de lucidité

## 3.4 Produire — Carte des ressources

### 3.4.1 Livrable principal

### 3.4.2 Tableau d’inventaire

### 3.4.3 Exemple appliqué au Grand Dossier COVID

## 3.5 Retenir — Ce qu’une carte permet et ne permet pas

## 3.6 Référencer — Références mobilisées

## 3.7 Livrable du chapitre

## 3.8 Transition vers le chapitre 4
```

---

## 10. Checklist de validation

Une fois un chapitre corrigé, vérifier :

- [ ] Le chapitre possède un seul titre `#`.
- [ ] Le titre de chapitre suit le format `# XX. Titre`.
- [ ] Aucune section interne n’utilise `#`.
- [ ] Toutes les sections principales utilisent `## XX.Y`.
- [ ] Toutes les sous-sections utilisent `### XX.Y.Z`.
- [ ] Aucun titre n’utilise `####`.
- [ ] Aucune section interne ne redémarre à `1.`, sauf si elle est une liste ordinaire.
- [ ] Les métadonnées UCKK sont en gras, non en heading.
- [ ] Les exemples, étapes et notes sont en gras ou en listes, non en titres de niveau 4.
- [ ] La transition finale est numérotée.
- [ ] Le livrable du chapitre est numéroté.
- [ ] La table des matières générée automatiquement serait lisible.

---

## 11. Prompt de référence pour corriger un chapitre

Utiliser ce prompt pour demander à une IA de corriger un chapitre.

```text
Corrige la structure Markdown du chapitre suivant selon le standard du manuscrit.

Règles obligatoires :
1. Un seul H1 par chapitre : `# XX. Titre du chapitre`.
2. Les sections principales sont en H2 et numérotées `XX.Y`.
3. Les sous-sections sont en H3 et numérotées `XX.Y.Z`.
4. Ne jamais dépasser trois niveaux de numérotation.
5. Convertir tous les H4 en texte gras ou en listes.
6. Convertir les métadonnées comme `Fonction UCKK dominante` et `Transformation du chapitre` en lignes en gras, non en titres.
7. Ne pas modifier le fond du contenu sauf pour harmoniser les titres.
8. Ne pas supprimer les exemples, tableaux, exercices ou transitions.
9. Si une section interne redémarre à `1.`, la renuméroter avec le numéro du chapitre.
10. Respecter la logique UCKK : Situer, Distinguer, Transformer, Produire, Retenir, Référencer, Livrable, Transition.

Retourne le chapitre corrigé en Markdown.
```

---

## 12. Décision finale

Le standard officiel du manuscrit est :

> `# XX. Chapitre`  
> `## XX.Y Section`  
> `### XX.Y.Z Sous-section`

Tout le reste — exemples, étapes, notes, variantes, consignes, points de vigilance — doit rester hors hiérarchie numérotée, sous forme de gras, listes ou encadrés.

Cette décision protège la lisibilité du livre, facilite la génération automatique d’une table des matières, et donne à l’IA une règle stable pour prolonger ou corriger le manuscrit.
