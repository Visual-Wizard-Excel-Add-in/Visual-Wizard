import { render, screen, fireEvent } from "@testing-library/react";

import Header from "../taskpane/components/Header";
import usePublicStore from "../taskpane/store/publicStore";

vi.mock("../taskpane/utils/store", () => ({
  default: vi.fn(),
}));

describe("Header", () => {
  const mockSetCategory = vi.fn();
  const mockSetOpenTab = vi.fn();

  beforeEach(() => {
    usePublicStore.mockReturnValue({
      setCategory: mockSetCategory,
      setOpenTab: mockSetOpenTab,
    });

    render(<Header />);
  });

  it("should render all categories", () => {
    expect(screen.getByRole("tab", { name: "수식" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "서식" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "매크로" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "유효성" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "공유하기" })).toBeInTheDocument();
  });

  it("should call setCategory and setOpenTab on category click", () => {
    const formulaTab = screen.getByRole("tab", { name: "수식" });

    fireEvent.click(formulaTab);

    expect(mockSetCategory).toHaveBeenCalledWith("Formula");
    expect(mockSetOpenTab).toHaveBeenCalledWith([]);
  });

  it("should call setCategory with correct value when different category are clicked", () => {
    const styleTab = screen.getByRole("tab", { name: "서식" });

    fireEvent.click(styleTab);

    expect(mockSetCategory).toHaveBeenCalledWith("Style");

    const macroTab = screen.getByRole("tab", { name: "매크로" });

    fireEvent.click(macroTab);

    expect(mockSetCategory).toHaveBeenCalledWith("Macro");
  });
});
