Download the 3 files the unrealpak.exe is redistributle from Unreal Engine then run the Run_Mod_List_With_Report.bat file to check out the program.

Playlist Original Load Order Feature ✅
New Functionality:
Playlist Creation: When creating a new playlist, it now stores the original load orders for each mod in originalLoadOrders property
Playlist Overwrite: When overwriting an existing playlist, it updates the original load orders
Playlist Editor: When editing a playlist, it captures the current load orders as the new original
Playlist Apply: When applying a playlist, it restores the original load orders that were stored when that playlist was created
How It Works:
When you create a playlist:

Captures current load orders of all enabled mods
Stores them in playlist.originalLoadOrders[FolderKey]
Preserves these values as the "original" state for that playlist
When you apply a playlist:

Uses stored original load orders instead of current mod load orders
Restores each mod to the exact load order it had when that playlist was created
Maintains playlist independence - each playlist has its own original state
When you edit a playlist:

Updates the original load orders to match the current state
Future applications will use the newly saved original values
Benefits:
Playlist Independence: Each playlist maintains its own original load order
No Cross-Contamination: Changing load orders in one playlist doesn't affect others
Reliable Restoration: Always restores to the exact state when playlist was created
Backward Compatible: Existing playlists will work (they'll just use current load orders as fallback)
Now when you switch between playlists, each one will restore to its original load order that was set when that specific playlist was created!
