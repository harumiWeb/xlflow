# xlflow MVP Quality Hardening Spec

## Goal

Prepare the post-MVP phase by standardizing real Excel COM verification, adding regression coverage for known round-trip risks, and aligning local workflow docs with actual repository practice.

## Active Slice

This planning session covers:

- defining the quality-hardening scope in `docs/specs/mvp-quality-hardening.md`
- writing the execution plan in `docs/superpowers/plans/2026-04-29-mvp-quality-hardening.md`
- refreshing task tracking so later agents can continue without rediscovering context

## Next Implementation Targets

1. document verified class/form/document-module round-trip behavior in repository docs
2. strengthen script regression tests around workbook-source transforms
3. add one obvious local verification entry point for fast automated checks
4. keep `tmp_workspaces` E2E verification as the required real-workbook path
