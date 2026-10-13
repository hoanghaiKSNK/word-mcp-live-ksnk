"""Switch only command/cwd, preserving all other parsed Codex settings."""
import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import tempfile
import tomllib
from datetime import datetime
from pathlib import Path


def switch(config, command, cwd, backup_dir):
    assert Path(command).is_file(), 'MCP executable not found'
    assert Path(cwd).is_dir(), 'MCP working directory not found'
    original=config.read_bytes();text=original.decode('utf-8');before=tomllib.loads(text)
    match=re.search(r'(?ms)^\[mcp_servers\.word_live\]\r?\n.*?(?=^\[|\Z)',text)
    assert match, 'Existing word_live entry required'
    block=match.group()
    for key,value in [('command',command),('cwd',cwd)]:
        replacement=key+' = '+json.dumps(value)
        block,count=re.subn(r'(?m)^'+key+r'\s*=.*$',lambda _:replacement,block)
        assert count==1, f'Expected exactly one {key} entry'
    updated=text[:match.start()]+block+text[match.end():]
    expected=copy.deepcopy(before)
    expected['mcp_servers']['word_live']['command']=command
    expected['mcp_servers']['word_live']['cwd']=cwd
    assert tomllib.loads(updated)==expected
    backup_dir.mkdir(parents=True,exist_ok=True)
    backup=backup_dir/f'config.before-word-switch-{datetime.now():%Y%m%d-%H%M%S-%f}.toml'
    shutil.copy2(config,backup)
    assert config.read_bytes()==original, 'Config changed concurrently; refusing overwrite'
    fd,temporary=tempfile.mkstemp(prefix='.word-mcp-config-',dir=config.parent)
    try:
        with os.fdopen(fd,'wb') as output:
            output.write(updated.encode('utf-8'));output.flush();os.fsync(output.fileno())
        os.replace(temporary,config)
    finally:Path(temporary).unlink(missing_ok=True)
    assert tomllib.loads(config.read_text(encoding='utf-8'))==expected
    return {'command':command,'cwd':cwd,'old_command':before['mcp_servers']['word_live']['command'],
            'old_cwd':before['mcp_servers']['word_live'].get('cwd'),'config_backup':str(backup),
            'config_before_sha256':hashlib.sha256(original).hexdigest(),
            'config_after_sha256':hashlib.sha256(config.read_bytes()).hexdigest(),
            'other_mcp_entries_unchanged':len(before['mcp_servers'])-1}


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--config',type=Path,required=True)
    parser.add_argument('--command',required=True);parser.add_argument('--cwd',required=True)
    parser.add_argument('--backups',type=Path,required=True);parser.add_argument('--receipt',type=Path,required=True)
    args=parser.parse_args()
    receipt=switch(args.config,args.command,args.cwd,args.backups)
    args.receipt.write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'updated_server':'word_live','other_mcp_entries_unchanged':receipt['other_mcp_entries_unchanged'],
                      'backup':receipt['config_backup']},ensure_ascii=False))
