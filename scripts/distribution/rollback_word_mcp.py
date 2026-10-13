"""Restore only the old Word MCP entry; retain other current Codex settings."""
import argparse
import json
from pathlib import Path
from update_mcp_config import switch

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--receipt',type=Path,required=True)
    parser.add_argument('--config',type=Path,required=True)
    args=parser.parse_args()
    receipt=json.loads(args.receipt.read_text(encoding='utf-8'))
    if not receipt.get('old_cwd'):
        raise SystemExit('Old cwd missing; inspect the original config backup before restoring.')
    result=switch(args.config,receipt['old_command'],receipt['old_cwd'],args.receipt.parent/'private-config-backups')
    print(json.dumps({'restored_server':'word_live','other_entries_preserved':result['other_mcp_entries_unchanged']}))
