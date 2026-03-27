import flet as ft
from adiantamento import AdiantamentoView


def main(page: ft.Page):
    page.title = "Perfetto Consolidador de Adiantamentos"
    page.window_width = 1024
    page.window_height = 768
    page.bgcolor = ft.Colors.BLUE_GREY_200
    page.padding = 0

    # ── AppBar ──────────────────────────────────────────────
    page.appbar = ft.AppBar(
        leading=ft.Icon(
            ft.Icons.KEYBOARD_COMMAND_KEY,
            color=ft.Colors.WHITE
        ),
        leading_width=60,
        title=ft.Text(
            "Perfetto Consolidador de Adiantamentos",
            color=ft.Colors.WHITE,
            weight=ft.FontWeight.BOLD,
        ),
        bgcolor=ft.Colors.BLUE_700,
    )

    # ── Ação do botão WhatsApp ───────────────────────────────
    def abrir_whatsapp(e):
        page.launch_url("https://wa.me/5549988369338")

    # ── Footer ──────────────────────────────────────────────
    footer = ft.Container(
        content=ft.Row(
            controls=[
                ft.TextButton(
                    content=ft.Row(
                        controls=[
                            ft.Icon(
                                ft.Icons.CHAT,
                                color=ft.Colors.GREEN_600,
                                size=20,
                            ),
                            ft.Text(
                                "Suporte",
                                color=ft.Colors.GREEN_600,
                                weight=ft.FontWeight.BOLD,
                            ),
                        ],
                        spacing=6,
                    ),
                    on_click=abrir_whatsapp,
                )
            ],
            alignment=ft.MainAxisAlignment.END,
        ),
        bgcolor=ft.Colors.GREY_200,
        padding=ft.padding.symmetric(horizontal=16, vertical=6),
        border=ft.border.only(top=ft.BorderSide(1, ft.Colors.GREY_400)),
    )

    # ── Conteúdo principal ───────────────────────────────────
    view = AdiantamentoView(page)

    # FilePicker precisa ser registrado no overlay da página
    page.overlay.append(view.file_picker)

    page.add(
        ft.Column(
            controls=[
                ft.Container(
                    content=view.build(),
                    expand=True,
                    padding=16,
                ),
                footer,
            ],
            expand=True,
            spacing=0,
        )
    )

    page.update()


ft.app(target=main)
