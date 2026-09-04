# /arena-status — Arena availability check

Run `arena_status` (with `live_check: true` only if the user asks to open
Arena). Report: platform, whether `Arena.Application` is registered, allowed
model roots, and the live-check result. If Arena is missing, stop and tell
the user a Windows host with licensed Arena is required for COM tools —
only `list_arena_models`, compound-file and results tools work without it.
