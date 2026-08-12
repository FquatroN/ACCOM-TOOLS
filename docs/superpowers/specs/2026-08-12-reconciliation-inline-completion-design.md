# Reconciliation Inline Completion Design

## Goal

Simplify reconciliation completion by replacing the confirmation popup with an inline comment field in the Current reconciliation panel, while preserving the mandatory-comment rule for non-zero differences. Make locked records more compact by reducing the vertical space between each source label and its details.

## Completion interface

For every active, started reconciliation, the Current reconciliation panel shows a **Completion comment** textarea directly above the completion button.

- The textarea is always visible while the reconciliation can be completed.
- When the difference is zero, the comment is optional and the button reads **Complete reconciliation**.
- When the difference is non-zero, the comment is marked required, the button reads **Force complete**, and the button remains disabled until the textarea contains non-whitespace text.
- Clicking the enabled button submits the completion immediately. No completion or force-completion popup is shown.
- Completed reconciliations keep their existing completion summary and do not show the editable textarea.

The required state must be understandable without relying only on color. The field label or helper text will state that a comment is required when the difference is non-zero.

## Draft comment state

The browser keeps the completion comment as a draft associated with the active reconciliation. Re-rendering the Current reconciliation panel after adding or removing a locked record must not discard the draft.

The draft is cleared when:

- completion succeeds;
- a different reconciliation is opened;
- the current reconciliation is deleted; or
- the user starts a new reconciliation.

If adding or removing records changes the difference between zero and non-zero, the same draft remains visible and the required state and button availability update from the new difference.

## Submission and validation

The inline button uses the existing reconciliation action flow:

- difference zero: submit `complete`, including the optional comment if entered;
- difference non-zero: submit `force_complete` with the mandatory trimmed comment.

Client-side validation prevents an empty or whitespace-only force-completion comment. Existing API and database validation remains unchanged as a second layer of protection. Action failures continue to use the page's existing reconciliation error presentation, and the draft comment remains available so the user can retry.

## Popup removal

Remove the completion modal markup, cached element references, event listeners, and modal-specific render/open/close/confirm functions. Other reconciliation dialogs and lifecycle actions are outside this change.

## Compact locked records

Keep the existing source, amount, remove action, and detail content. Reduce only the vertical spacing between the source row and the detail line below it:

- use a smaller row gap;
- tighten the detail line height;
- avoid additional top margin on the detail line.

The source label and details must remain readable, and the compact styling must stay scoped to locked records in the Current reconciliation panel.

## Verification

Automated tests will verify:

- the completion modal is no longer rendered or wired;
- the textarea appears for a started reconciliation at both zero and non-zero differences;
- zero difference permits completion with an empty comment;
- non-zero difference disables Force complete for empty or whitespace-only input;
- entering a valid comment enables Force complete and submits it;
- the comment draft survives Current reconciliation re-renders and is cleared at the specified lifecycle boundaries;
- completed reconciliations do not show the editable completion controls;
- locked-record detail styling uses the tighter scoped spacing.

Manual browser verification will cover both completion paths and visually confirm the compact locked-record layout.

## Out of scope

- Changes to reconciliation arithmetic or source rules.
- Changes to database schemas or RPC signatures.
- Changes to comment requirements for delete or reopen actions.
- Redesigning the rest of the Current reconciliation panel.
