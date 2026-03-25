package core

import (
	"strings"
	"testing"

	"github.com/duynguyendang/docxgo/v3/domain"
	"github.com/duynguyendang/docxgo/v3/internal/manager"
)

func TestParagraph_AddTrackedInsertion(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	err := para.AddTrackedInsertion("Claude", "2025-01-01T00:00:00Z", func(run domain.Run) {
		_ = run.SetText("inserted text")
		_ = run.SetBold(true)
	})

	if err != nil {
		t.Fatalf("AddTrackedInsertion failed: %v", err)
	}

	// Check tracked changes are stored
	p := para.(*paragraph)
	changes := p.TrackedChanges()
	if len(changes) != 1 {
		t.Fatalf("expected 1 tracked change, got %d", len(changes))
	}

	change := changes[0]
	if change.Author() != "Claude" {
		t.Errorf("author = %q, want %q", change.Author(), "Claude")
	}
	if change.Date() != "2025-01-01T00:00:00Z" {
		t.Errorf("date = %q, want %q", change.Date(), "2025-01-01T00:00:00Z")
	}
	if change.Type() != domain.TrackedChangeInsertion {
		t.Errorf("type = %v, want %v", change.Type(), domain.TrackedChangeInsertion)
	}
	runs := change.Runs()
	if len(runs) != 1 {
		t.Fatalf("expected 1 run in tracked change, got %d", len(runs))
	}
	if runs[0].Text() != "inserted text" {
		t.Errorf("run text = %q, want %q", runs[0].Text(), "inserted text")
	}
	if !runs[0].Bold() {
		t.Error("run should be bold")
	}
}

func TestParagraph_AddTrackedDeletion(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	err := para.AddTrackedDeletion("Claude", "", func(run domain.Run) {
		_ = run.SetText("deleted text")
	})

	if err != nil {
		t.Fatalf("AddTrackedDeletion failed: %v", err)
	}

	p := para.(*paragraph)
	changes := p.TrackedChanges()
	if len(changes) != 1 {
		t.Fatalf("expected 1 tracked change, got %d", len(changes))
	}

	change := changes[0]
	if change.Type() != domain.TrackedChangeDeletion {
		t.Errorf("type = %v, want %v", change.Type(), domain.TrackedChangeDeletion)
	}
}

func TestParagraph_AddTrackedInsertion_EmptyAuthor(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	err := para.AddTrackedInsertion("", "", func(run domain.Run) {
		_ = run.SetText("text")
	})

	if err == nil {
		t.Error("expected error for empty author")
	}
}

func TestParagraph_AddComment(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	comment, err := para.AddComment("Claude", "C", "This is a comment")
	if err != nil {
		t.Fatalf("AddComment failed: %v", err)
	}

	if comment.Author() != "Claude" {
		t.Errorf("author = %q, want %q", comment.Author(), "Claude")
	}
	if comment.Initials() != "C" {
		t.Errorf("initials = %q, want %q", comment.Initials(), "C")
	}
	if comment.Text() != "This is a comment" {
		t.Errorf("text = %q, want %q", comment.Text(), "This is a comment")
	}

	// Check comments are accessible via paragraph
	comments := para.Comments()
	if len(comments) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(comments))
	}
}

func TestParagraph_AddComment_EmptyAuthor(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	_, err := para.AddComment("", "", "text")
	if err == nil {
		t.Error("expected error for empty author")
	}
}

func TestParagraph_MultipleTrackedChanges(t *testing.T) {
	idGen := manager.NewIDGenerator()
	relManager := manager.NewRelationshipManager(idGen)
	mediaManager := manager.NewMediaManager(idGen)
	para := NewParagraph("para1", idGen, relManager, mediaManager)

	_ = para.AddTrackedInsertion("Claude", "", func(run domain.Run) {
		_ = run.SetText("first")
	})
	_ = para.AddTrackedDeletion("Claude", "", func(run domain.Run) {
		_ = run.SetText("second")
	})
	_ = para.AddTrackedInsertion("Jane", "", func(run domain.Run) {
		_ = run.SetText("third")
	})

	p := para.(*paragraph)
	changes := p.TrackedChanges()
	if len(changes) != 3 {
		t.Fatalf("expected 3 tracked changes, got %d", len(changes))
	}

	// Verify IDs are unique
	ids := make(map[string]bool)
	for _, ch := range changes {
		if ids[ch.ID()] {
			t.Errorf("duplicate tracked change ID: %s", ch.ID())
		}
		ids[ch.ID()] = true
	}
}

func TestCurrentISOTimestamp(t *testing.T) {
	ts := currentISOTimestamp()
	if !strings.Contains(ts, "T") || !strings.Contains(ts, "Z") {
		t.Errorf("expected ISO 8601 format, got %q", ts)
	}
}
