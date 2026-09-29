"""Word sesijas īpašumtiesību testi (bez reāla Word, ar COM aizvietotājiem).

Galvenā prasība: `DispatchEx -> Dispatch` fallback laikā NEKAD nedrīkst aizvērt
lietotāja jau atvērto Word sesiju.
"""
import os
import types

import pytest

from docmerge.engines import word_com
from docmerge.testing.fakes import FakeWord, fake_com_environment, word_unavailable_client


def _force_dispatch_only(monkeypatch):
    monkeypatch.setenv(word_com.DISPATCH_METHODS_ENV,'Dispatch')
    monkeypatch.setattr(word_com,'WORD_DISPATCH_RETRY_DELAY',0)


def test_dispatch_to_existing_word_is_not_owned(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord()
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids={1111}):
        session=word_com.open_word_session(client)
        assert session.owned is False
        assert session.safe_to_quit is False
        assert session.method=='Dispatch'
        assert session.close() is False
        assert word.quit_calls==0
        assert word.Documents.Count==0  # LIETOTĀJA sesijā nekas netika atstāts


def test_dispatch_without_existing_word_is_owned(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord()
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids=set()):
        session=word_com.open_word_session(client)
        assert session.owned is True
        assert session.close() is True
        assert word.quit_calls==1


def test_unknown_word_state_is_treated_as_user_session(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord()
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids=None):
        session=word_com.open_word_session(client)
        assert session.owned is False,'Nezināms stāvoklis nedrīkst ļaut Quit (datu drošība)'


def test_dispatchex_is_always_owned(monkeypatch):
    monkeypatch.setattr(word_com,'WORD_DISPATCH_RETRY_DELAY',0)
    word=FakeWord()
    client=types.SimpleNamespace(DispatchEx=lambda name:word,Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids={2222}):
        session=word_com.open_word_session(client)
        assert session.owned is True and session.method=='DispatchEx'
        assert session.close() is True
        assert word.quit_calls==1


def test_user_session_display_alerts_restored(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord(); word.DisplayAlerts=-1
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids={3333}):
        session=word_com.open_word_session(client)
        assert word.DisplayAlerts==0
        session.close()
        assert word.DisplayAlerts==-1,'DisplayAlerts netika atjaunots lietotāja sesijai'


def test_user_session_visible_is_not_changed(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord(); word.Visible=True
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids={4444}):
        word_com.open_word_session(client)
        assert word.Visible is True,'Lietotāja sesijai Visible nedrīkst mainīt'


def test_own_session_is_hidden(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord(); word.Visible=True
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids=set()):
        word_com.open_word_session(client)
        assert word.Visible is False


def test_word_unavailable_raises_with_full_text(monkeypatch):
    monkeypatch.setattr(word_com,'WORD_DISPATCH_RETRY_DELAY',0)
    failing=word_unavailable_client()
    with fake_com_environment(client=failing,winword_pids=set()):
        with pytest.raises(Exception) as info:
            word_com.open_word_session(failing)
        assert 'Server execution failed' in str(info.value)


def test_configured_dispatch_methods_env_override(monkeypatch):
    monkeypatch.delenv(word_com.DISPATCH_METHODS_ENV,raising=False)
    assert word_com.configured_dispatch_methods()==word_com.WORD_DISPATCH_METHODS
    monkeypatch.setenv(word_com.DISPATCH_METHODS_ENV,'Dispatch')
    assert word_com.configured_dispatch_methods()==('Dispatch',)
    monkeypatch.setenv(word_com.DISPATCH_METHODS_ENV,'Dispatch,DispatchEx')
    assert word_com.configured_dispatch_methods()==('Dispatch','DispatchEx')


def test_word_is_running_without_probe_returns_none(monkeypatch):
    """Neziņa par WINWORD stāvokli nozīmē 'nav mūsu instance' (datu drošība)."""
    monkeypatch.setattr(word_com,'word_process_ids',lambda:None)
    assert word_com.word_is_running() is None
    assert word_com.word_is_running(set()) is False
    assert word_com.word_is_running({7}) is True


def test_session_close_is_idempotent(monkeypatch):
    _force_dispatch_only(monkeypatch)
    word=FakeWord()
    client=types.SimpleNamespace(Dispatch=lambda name:word)
    with fake_com_environment(word=word,client=client,winword_pids=set()):
        session=word_com.open_word_session(client)
        assert session.close() is True
        assert session.close() is False
        assert word.quit_calls==1
