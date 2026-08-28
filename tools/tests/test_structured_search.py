import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
MODULE_PATH = ROOT / "CEDLib.lib" / "UIClasses" / "structured_search.py"
SPEC = importlib.util.spec_from_file_location("ced_structured_search", str(MODULE_PATH))
SEARCH = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SEARCH)


class Item(object):
    def __init__(self, panel, poles, search_text):
        self.panel = panel
        self.poles = poles
        self.search_text = search_text


def definitions():
    return [
        SEARCH.SearchFilterDefinition(
            "panel",
            "Panel",
            matcher=lambda item, value: item.panel.lower() == value.lower(),
        ),
        SEARCH.SearchFilterDefinition(
            "poles",
            "Poles",
            matcher=lambda item, value: item.poles == int(value),
        ),
        SEARCH.SearchFilterDefinition(
            "status",
            "Status",
            matcher=lambda item, value: value.lower() in item.search_text.lower(),
        ),
    ]


class StructuredSearchStateTests(unittest.TestCase):
    def setUp(self):
        self.state = SEARCH.StructuredSearchState(definitions())
        self.query_events = []
        self.command_events = []
        self.state.add_query_changed_handler(
            lambda sender, args: self.query_events.append(args.query)
        )
        self.state.add_command_changed_handler(
            lambda sender, args: self.command_events.append(args)
        )

    def choose(self, key):
        self.state.begin_command()
        return self.state.select_filter(key)

    def test_plain_free_text_is_live(self):
        self.state.set_input_text("rec")

        self.assertEqual("rec", self.state.query.FreeText)
        self.assertEqual((), self.state.query.Filters)
        self.assertEqual(1, len(self.query_events))

    def test_free_text_preserves_spaces_while_query_matching_trims_edges(self):
        self.state.set_input_text("receptacle ")

        self.assertEqual("receptacle ", self.state.free_text)
        self.assertEqual("receptacle", self.state.query.FreeText)

        self.state.set_input_text("receptacle panel")
        self.assertEqual("receptacle panel", self.state.free_text)
        self.assertEqual("receptacle panel", self.state.query.FreeText)

    def test_command_narrowing_does_not_change_results(self):
        self.state.set_input_text("rec")
        before = self.state.query

        self.state.set_input_text("rec /p")

        self.assertTrue(self.state.is_command_mode)
        self.assertEqual("p", self.state.command_text)
        self.assertEqual("rec /p", self.state.command_input_text)
        self.assertEqual(["panel", "poles"], [x.key for x in self.state.command_suggestions])
        self.assertEqual(before, self.state.query)
        self.assertEqual(1, len(self.query_events))

    def test_slash_requires_input_start_or_preceding_whitespace(self):
        self.state.set_input_text("rec/p")

        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec/p", self.state.free_text)

        self.state.set_input_text("rec /p")

        self.assertTrue(self.state.is_command_mode)
        self.assertEqual("rec", self.state.free_text)
        self.assertEqual("p", self.state.command_text)

    def test_escaping_slash_mode_keeps_the_full_input_as_free_text(self):
        self.state.set_input_text("rec /pa")

        self.assertTrue(self.state.cancel_command_as_literal())
        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec /pa", self.state.free_text)
        self.assertEqual("rec /pa", self.state.query.FreeText)

    def test_unmatched_slash_keyword_becomes_plain_text_and_stays_plain(self):
        self.state.set_input_text("rec /unknown")

        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec /unknown", self.state.free_text)

        self.state.set_input_text("rec /unknown more")

        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec /unknown more", self.state.free_text)

    def test_space_after_slash_fragment_returns_to_plain_text(self):
        self.state.set_input_text("rec /panel ")

        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec /panel ", self.state.free_text)

    def test_escape_cancels_command_without_altering_query(self):
        self.state.set_input_text("rec")
        before = self.state.query
        self.state.begin_command("pa")

        self.assertTrue(self.state.cancel_command())
        self.assertFalse(self.state.is_command_mode)
        self.assertEqual("rec", self.state.free_text)
        self.assertEqual(before, self.state.query)
        self.assertEqual(1, len(self.query_events))

    def test_selecting_filter_creates_empty_token_without_filtering(self):
        token = self.choose("panel")

        self.assertEqual("panel", token.key)
        self.assertEqual(1, len(self.state.tokens))
        self.assertEqual(0, len(self.state.query.Filters))
        self.assertEqual(0, len(self.query_events))
        self.assertEqual(0, self.state.active_token_index)

    def test_filter_value_is_live_and_editable(self):
        self.state.set_free_text("rec")
        self.choose("panel")
        self.state.set_active_token_value("A")

        self.assertEqual("rec", self.state.query.FreeText)
        self.assertEqual("A", self.state.query.Filters[0].Value)
        self.assertEqual(2, len(self.query_events))

        self.state.commit_active_token()
        self.state.edit_token(0)
        self.state.set_active_token_value("A1")
        self.assertEqual("A1", self.state.query.Filters[0].Value)
        self.assertEqual(3, len(self.query_events))

    def test_multiple_filters_use_and_semantics(self):
        self.choose("panel")
        self.state.set_active_token_value("A")
        self.state.commit_active_token()
        self.choose("poles")
        self.state.set_active_token_value("3")

        self.assertEqual(["panel", "poles"], [x.Key for x in self.state.query.Filters])
        self.assertEqual(["A", "3"], [x.Value for x in self.state.query.Filters])

        matcher = lambda item, text: text.lower() in item.search_text.lower()
        self.assertTrue(self.state.query.matches(Item("A", 3, "receptacle"), matcher))
        self.assertFalse(self.state.query.matches(Item("B", 3, "receptacle"), matcher))
        self.assertFalse(self.state.query.matches(Item("A", 2, "receptacle"), matcher))

    def test_selecting_an_existing_key_reactivates_instead_of_duplicating(self):
        first = self.choose("panel")
        self.state.set_active_token_value("A")
        self.state.commit_active_token()

        second = self.choose("panel")

        self.assertIs(first, second)
        self.assertEqual(1, len(self.state.tokens))
        self.assertEqual(0, self.state.active_token_index)
        self.assertEqual("A", self.state.query.Filters[0].Value)

    def test_multi_value_filters_are_or_grouped_by_key(self):
        multi_state = SEARCH.StructuredSearchState(
            [
                SEARCH.SearchFilterDefinition(
                    "panel",
                    "Panel",
                    allow_multiple=True,
                    matcher=lambda item, value: item.panel.lower() == value.lower(),
                ),
                SEARCH.SearchFilterDefinition(
                    "poles",
                    "Poles",
                    allow_multiple=True,
                    matcher=lambda item, value: item.poles == int(value),
                ),
            ]
        )

        multi_state.begin_command()
        multi_state.select_filter("panel")
        multi_state.set_active_token_value("A")
        multi_state.commit_active_token()
        multi_state.begin_command()
        multi_state.select_filter("panel")
        multi_state.set_active_token_value("B")
        multi_state.commit_active_token()
        multi_state.begin_command()
        multi_state.select_filter("poles")
        multi_state.set_active_token_value("3")

        query = multi_state.query
        self.assertEqual(["panel", "panel", "poles"], [item.Key for item in query.Filters])
        self.assertTrue(query.matches(Item("A", 3, "receptacle"), lambda item, text: True))
        self.assertTrue(query.matches(Item("B", 3, "receptacle"), lambda item, text: True))
        self.assertFalse(query.matches(Item("C", 3, "receptacle"), lambda item, text: True))
        self.assertFalse(query.matches(Item("A", 2, "receptacle"), lambda item, text: True))

    def test_tokens_are_grouped_by_definition_order(self):
        self.state.begin_command()
        self.state.select_filter("poles")
        self.state.set_active_token_value("3")
        self.state.commit_active_token()
        self.state.begin_command()
        self.state.select_filter("panel")
        self.state.set_active_token_value("A")

        self.assertEqual(["panel", "poles"], [x.key for x in self.state.tokens])

    def test_atomic_backspace_selects_then_removes_token(self):
        self.choose("panel")
        self.state.set_active_token_value("A")
        self.state.commit_active_token()
        self.choose("poles")
        self.state.set_active_token_value("3")
        self.state.commit_active_token()
        before = self.state.query

        self.assertTrue(self.state.backspace_at_input_start())
        self.assertEqual(1, self.state.selected_token_index)
        self.assertEqual(before, self.state.query)

        self.assertTrue(self.state.backspace_at_input_start())
        self.assertEqual(["panel"], [x.key for x in self.state.query.Filters])

    def test_delete_has_the_same_two_step_atomic_behavior(self):
        self.choose("panel")
        self.state.set_active_token_value("A")
        self.state.commit_active_token()
        self.choose("poles")
        self.state.set_active_token_value("3")
        self.state.commit_active_token()

        self.assertTrue(self.state.delete_at_input_start())
        self.assertEqual(1, self.state.selected_token_index)
        self.assertTrue(self.state.delete_at_input_start())
        self.assertEqual(["panel"], [x.key for x in self.state.query.Filters])

    def test_remove_filter_and_clear_reset_effective_query(self):
        self.choose("panel")
        self.state.set_active_token_value("A")
        self.state.commit_active_token()
        self.assertTrue(self.state.remove_token(0))
        self.assertTrue(self.state.query.is_empty)

        self.state.set_input_text("text")
        self.assertTrue(self.state.clear())
        self.assertTrue(self.state.query.is_empty)
        self.assertFalse(self.state.has_content)


if __name__ == "__main__":
    unittest.main()
