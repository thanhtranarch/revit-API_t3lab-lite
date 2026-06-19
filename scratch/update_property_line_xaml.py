# -*- coding: utf-8 -*-
import io

file_path = 'T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml'

with io.open(file_path, mode='r', encoding='utf-8') as f:
    content = f.read()

# Replacement 1: WindowChrome corner radius to 22
target_1 = """        <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                       ResizeBorderThickness="5"
                       GlassFrameThickness="0"
                       CornerRadius="8"
                       UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>"""

replacement_1 = """    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="22"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>"""

# Replacement 2: Wrap root Grid in Border
target_2 = """    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="64"/>    <!-- Modern Title bar -->
            <RowDefinition Height="*"/>     <!-- 2-Column Content -->
            <RowDefinition Height="Auto"/>  <!-- Bottom Action + Status bar -->
        </Grid.RowDefinitions>"""

replacement_2 = """    <!-- ═══ Outer Border: CornerRadius + ClipToBounds (Lumina standard) ═══ -->
    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>    <!-- Modern Title bar -->
                <RowDefinition Height="*"/>     <!-- 2-Column Content -->
                <RowDefinition Height="Auto"/>  <!-- Bottom Action + Status bar -->
            </Grid.RowDefinitions>"""

# Replacement 3: API Status block with Hyperlink
target_3 = """                            <TextBlock x:Name="txt_api_status"
                                       Text="No API key configured"
                                       Foreground="#EF4444"
                                       FontSize="11"
                                       FontWeight="Medium"/>"""

replacement_3 = """                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name="txt_api_status"
                                           Grid.Column="0"
                                           Text="No API key configured"
                                           Foreground="#EF4444"
                                           FontSize="11"
                                           FontWeight="Medium"
                                           VerticalAlignment="Center"/>
                                <TextBlock Grid.Column="1" FontSize="11" VerticalAlignment="Center">
                                    <Hyperlink Click="btn_get_api_key_Click" ToolTip="Register or manage Lightbox RE API keys">
                                        <Run Text="Get API Key" Foreground="#3B82F6" TextDecorations="Underline"/>
                                    </Hyperlink>
                                </TextBlock>
                            </Grid>"""

# Replacement 4: Close Border at the end of the file
target_4 = """        </Border></Grid>
</Window>"""

replacement_4 = """        </Border>
        <!-- Copyright added automatically -->
        <TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999" Grid.RowSpan="99" Grid.ColumnSpan="99"/>
        </Grid>
    </Border>
</Window>"""

# Perform replacements and check if they matched
replaced_count = 0

if target_1 in content:
    content = content.replace(target_1, replacement_1)
    print("Success: Replaced WindowChrome CornerRadius")
    replaced_count += 1
else:
    normalized_target_1 = target_1.replace("\r\n", "\n")
    normalized_content = content.replace("\r\n", "\n")
    if normalized_target_1 in normalized_content:
        normalized_content = normalized_content.replace(normalized_target_1, replacement_1.replace("\r\n", "\n"))
        content = normalized_content
        print("Success: Replaced WindowChrome CornerRadius (normalized newlines)")
        replaced_count += 1
    else:
        print("Error: Target 1 not found in content")

if target_2 in content:
    content = content.replace(target_2, replacement_2)
    print("Success: Replaced Root Grid wrapper")
    replaced_count += 1
else:
    normalized_target_2 = target_2.replace("\r\n", "\n")
    normalized_content = content.replace("\r\n", "\n")
    if normalized_target_2 in normalized_content:
        normalized_content = normalized_content.replace(normalized_target_2, replacement_2.replace("\r\n", "\n"))
        content = normalized_content
        print("Success: Replaced Root Grid wrapper (normalized newlines)")
        replaced_count += 1
    else:
        print("Error: Target 2 not found in content")

if target_3 in content:
    content = content.replace(target_3, replacement_3)
    print("Success: Replaced API Status block")
    replaced_count += 1
else:
    normalized_target_3 = target_3.replace("\r\n", "\n")
    normalized_content = content.replace("\r\n", "\n")
    if normalized_target_3 in normalized_content:
        normalized_content = normalized_content.replace(normalized_target_3, replacement_3.replace("\r\n", "\n"))
        content = normalized_content
        print("Success: Replaced API Status block (normalized newlines)")
        replaced_count += 1
    else:
        print("Error: Target 3 not found in content")

if target_4 in content:
    content = content.replace(target_4, replacement_4)
    print("Success: Replaced Closing Tag")
    replaced_count += 1
else:
    normalized_target_4 = target_4.replace("\r\n", "\n")
    normalized_content = content.replace("\r\n", "\n")
    if normalized_target_4 in normalized_content:
        normalized_content = normalized_content.replace(normalized_target_4, replacement_4.replace("\r\n", "\n"))
        content = normalized_content
        print("Success: Replaced Closing Tag (normalized newlines)")
        replaced_count += 1
    else:
        print("Error: Target 4 not found in content")

# Save if all replacements worked
if replaced_count == 4:
    with io.open(file_path, mode='w', encoding='utf-8') as f:
        f.write(content)
    print("All 4 changes written successfully to PropertyLine.xaml")
else:
    print("Failed to apply all 4 changes, not saving.")
