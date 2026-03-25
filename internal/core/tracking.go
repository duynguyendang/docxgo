package core

import (
	"time"

	"github.com/duynguyendang/docxgo/v3/domain"
)

// trackedInsertion implements a tracked insertion (<w:ins>).
type trackedInsertion struct {
	id     string
	author string
	date   string
	runs   []domain.Run
}

func (t *trackedInsertion) ID() string     { return t.id }
func (t *trackedInsertion) Author() string { return t.author }
func (t *trackedInsertion) Date() string   { return t.date }
func (t *trackedInsertion) Runs() []domain.Run {
	runs := make([]domain.Run, len(t.runs))
	copy(runs, t.runs)
	return runs
}
func (t *trackedInsertion) Type() domain.TrackedChangeType {
	return domain.TrackedChangeInsertion
}

// trackedDeletion implements a tracked deletion (<w:del>).
type trackedDeletion struct {
	id     string
	author string
	date   string
	runs   []domain.Run
}

func (t *trackedDeletion) ID() string     { return t.id }
func (t *trackedDeletion) Author() string { return t.author }
func (t *trackedDeletion) Date() string   { return t.date }
func (t *trackedDeletion) Runs() []domain.Run {
	runs := make([]domain.Run, len(t.runs))
	copy(runs, t.runs)
	return runs
}
func (t *trackedDeletion) Type() domain.TrackedChangeType {
	return domain.TrackedChangeDeletion
}

// trackedChangeRun is the internal interface for tracked change elements.
type trackedChangeRun interface {
	ID() string
	Author() string
	Date() string
	Type() domain.TrackedChangeType
	Runs() []domain.Run
}

// docxComment implements the domain.Comment interface.
type docxComment struct {
	id       string
	author   string
	date     string
	initials string
	text     string
	paraID   string
}

func (c *docxComment) ID() string       { return c.id }
func (c *docxComment) Author() string   { return c.author }
func (c *docxComment) Date() string     { return c.date }
func (c *docxComment) Initials() string { return c.initials }
func (c *docxComment) Text() string     { return c.text }
func (c *docxComment) ParaID() string   { return c.paraID }

// currentISOTimestamp returns the current time in ISO 8601 format.
func currentISOTimestamp() string {
	return time.Now().UTC().Format("2006-01-02T15:04:05Z")
}
