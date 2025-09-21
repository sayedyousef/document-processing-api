-              # Accept tracked changes
       415 -              print("\nAccepting tracked changes...")
       414 +              # Check for tracked changes
       415 +              print("\nChecking for tracked changes...")
       416 +              has_tracked_changes = False
       417                try:
       418 -                  self.doc.AcceptAllRevisions()
       419 -                  print("✓ All revisions accepted")
       420 -              except:
       421 -                  print("⚠ No tracked changes or unable to accept")
       418 +                  # Check if track changes is enabled or if there are existing revisions
       419 +                  if self.doc.TrackRevisions:
       420 +                      has_tracked_changes = True
       421 +                      print("❌ Track Changes is ENABLED in this document")
       422 +                  elif self.doc.Revisions.Count > 0:
       423 +                      has_tracked_changes = True
       424 +                      print(f"❌ Document has {self.doc.Revisions.Count} tracked changes")
       425 +
       426 +                  if has_tracked_changes:
       427 +                      error_msg = f"Document '{docx_path.name}' has tracked changes and cannot be processed. Please accept all changes and
           + disable tracking before processing."
       428 +                      print(f"\n{error_msg}")
       429 +                      return {
       430 +                          'error': error_msg,
       431 +                          'has_tracked_changes': True,
       432 +                          'file_name': docx_path.name
       433 +                      }
       434 +                  else:
       435 +                      print("✓ No tracked changes detected")
       436 +              except Exception as e:
       437 +                  print(f"⚠ Warning: Could not check tracked changes: {e}")