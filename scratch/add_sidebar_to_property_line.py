# -*- coding: utf-8 -*-
import io

file_path = 'T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml'

with io.open(file_path, mode='r', encoding='utf-8') as f:
    content = f.read()

# Replacement 1: Add ColumnDefinitions and Sidebar Border
target_1 = """        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>    <!-- Modern Title bar -->
                <RowDefinition Height="*"/>     <!-- 2-Column Content -->
                <RowDefinition Height="Auto"/>  <!-- Bottom Action + Status bar -->
            </Grid.RowDefinitions>"""

replacement_1 = """        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="66"/>    <!-- Col 0: Left Sidebar Icon Rail -->
                <ColumnDefinition Width="*"/>     <!-- Col 1: Main content area -->
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>    <!-- Modern Title bar -->
                <RowDefinition Height="*"/>     <!-- 2-Column Content -->
                <RowDefinition Height="Auto"/>  <!-- Bottom Action + Status bar -->
            </Grid.RowDefinitions>

            <!-- COL 0: LEFT SIDEBAR ICON RAIL (spans all rows) -->
            <Border Grid.Column="0" Grid.RowSpan="3"
                    Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,0,1,0">
                <Grid Margin="0,16,0,16">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Logo -->
                        <RowDefinition Height="Auto"/> <!-- Active tab icon -->
                        <RowDefinition Height="*"/>    <!-- Spacer -->
                    </Grid.RowDefinitions>

                    <!-- Logo: Map pin logo -->
                    <Border Grid.Row="0" Width="42" Height="42" CornerRadius="13"
                            Background="White" BorderBrush="#DCDCE0" BorderThickness="1"
                            HorizontalAlignment="Center" Margin="0,0,0,20">
                        <Path Data="M12 2 C8 2 4 6 4 10 C4 16 12 22 12 22 C12 22 20 16 20 10 C20 6 16 2 12 2 Z M12 13 A3 3 0 1 1 12 7 A3 3 0 1 1 12 13 Z"
                              Stroke="#18181B" StrokeThickness="1.6" StrokeLineJoin="Round"
                              StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Width="18" Height="18" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>

                    <!-- Tab: Property Search (Active/Selected toggle-like look) -->
                    <Border Grid.Row="1" Width="42" Height="42" CornerRadius="13" Background="#18181B" HorizontalAlignment="Center">
                        <Path Data="M4 5 h16 v14 h-16 z M4 10 h16 M10 10 v9"
                              Stroke="#FFFFFF" StrokeThickness="1.8" StrokeLineJoin="Round"
                              Width="17" Height="17" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                </Grid>
            </Border>"""

# Replacement 2: Move Title Bar to Col 1
target_2 = """        <!-- ═══════════════════════════════════════════════════════════ MODERN TITLE BAR -->
        <Grid Grid.Row="0" Background="White" Height="64">"""

replacement_2 = """        <!-- ═══════════════════════════════════════════════════════════ MODERN TITLE BAR -->
        <Grid Grid.Column="1" Grid.Row="0" Background="White" Height="64">"""

# Replacement 3: Move Content to Col 1
target_3 = """        <!-- ═══════════════════════════════════════════════════════════ CONTENT (2-COLUMN) -->
        <Grid Grid.Row="1" Background="White">"""

replacement_3 = """        <!-- ═══════════════════════════════════════════════════════════ CONTENT (2-COLUMN) -->
        <Grid Grid.Column="1" Grid.Row="1" Background="White">"""

# Replacement 4: Move Bottom Bar to Col 1
target_4 = """        <!-- ═══════════════════════════════════════════════════════════ COMBINED BOTTOM BAR -->
        <Border Grid.Row="2" Background="#F8FAFC\""""

replacement_4 = """        <!-- ═══════════════════════════════════════════════════════════ COMBINED BOTTOM BAR -->
        <Border Grid.Column="1" Grid.Row="2" Background="#F8FAFC\""""

replaced = 0

# Check targets and perform replacements (supporting normalization for CRLF newlines)
for tgt, rep, name in [
    (target_1, replacement_1, "Sidebar Grid Setup"),
    (target_2, replacement_2, "Title Bar Grid Column"),
    (target_3, replacement_3, "Content Grid Column"),
    (target_4, replacement_4, "Bottom Bar Grid Column")
]:
    if tgt in content:
        content = content.replace(tgt, rep)
        print("Success: Applied {}".format(name))
        replaced += 1
    else:
        norm_tgt = tgt.replace("\r\n", "\n")
        norm_content = content.replace("\r\n", "\n")
        if norm_tgt in norm_content:
            norm_content = norm_content.replace(norm_tgt, rep.replace("\r\n", "\n"))
            content = norm_content
            print("Success: Applied {} (normalized newlines)".format(name))
            replaced += 1
        else:
            print("Error: Target for {} not found".format(name))

if replaced == 4:
    with io.open(file_path, mode='w', encoding='utf-8') as f:
        f.write(content)
    print("All 4 sidebar changes applied successfully to PropertyLine.xaml")
else:
    print("Failed to apply all changes, not saving.")
